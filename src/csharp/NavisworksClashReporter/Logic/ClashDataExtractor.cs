using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using HtmlAgilityPack;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;

namespace NavisworksClashReporter.Logic
{
    public class ClashInfo
    {
        public string Title { get; set; }
        public string MainFolder { get; set; }
        public bool IsResolved { get; set; }
        public string DateFound { get; set; }
        public string ComputedStatus { get; set; }
        public string ImagePath { get; set; }
    }

    public class ClashDataExtractor
    {
        private string NormalizeTitle(string title)
        {
            if (string.IsNullOrEmpty(title)) return "";
            return System.Text.RegularExpressions.Regex.Replace(title, @"\s+", " ").Trim();
        }

        public List<string> GetBuildingsFromApi()
        {
            var buildings = new List<string>();
            try
            {
                foreach (var item in Application.ActiveDocument.SavedViewpoints.RootItem.Children)
                {
                    if (item is FolderItem folder)
                    {
                        buildings.Add(folder.DisplayName.Trim());
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
            }
            return buildings;
        }

        public List<string> GetBuildingsFromXml(string xmlPath)
        {
            var buildings = new List<string>();
            try
            {
                if (File.Exists(xmlPath))
                {
                    XDocument doc = XDocument.Load(xmlPath);
                    var viewpointsNode = doc.Descendants("viewpoints").FirstOrDefault();
                    if (viewpointsNode != null)
                    {
                        foreach (var child in viewpointsNode.Elements())
                        {
                            if (child.Name.LocalName == "viewfolder")
                            {
                                string name = child.Attribute("name")?.Value.Trim() ?? "";
                                if (!string.IsNullOrEmpty(name)) buildings.Add(name);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
            }
            return buildings;
        }

        public Dictionary<string, ClashInfo> ExtractFromXmlHtml(string xmlPath, string htmlFolder, string targetBuilding = null)
        {
            var viewData = new Dictionary<string, ClashInfo>();
            var looseViewData = new Dictionary<string, string>();

            try
            {
                // Parse XML
                XDocument doc = XDocument.Load(xmlPath);
                
                void Traverse(XElement element, List<string> currentPath)
                {
                    foreach (var child in element.Elements())
                    {
                        if (child.Name.LocalName == "viewfolder")
                        {
                            string folderName = child.Attribute("name")?.Value.Trim() ?? "";
                            
                            if (currentPath.Count == 0 && !string.IsNullOrEmpty(targetBuilding))
                            {
                                if (folderName != targetBuilding) continue;
                            }

                            var newPath = new List<string>(currentPath) { folderName };
                            Traverse(child, newPath);
                        }
                        else if (child.Name.LocalName == "view")
                        {
                            // Ignore views that are not inside any folder (root-level views)
                            if (currentPath.Count == 0) continue;

                            string viewName = child.Attribute("name")?.Value.Trim() ?? "";
                            viewName = NormalizeTitle(viewName);
                            bool isResolved = currentPath.Any(f => f.ToUpper() == "RESOLVED");
                            
                            string dateFound = "";
                            for (int i = currentPath.Count - 1; i >= 0; i--)
                            {
                                if (currentPath[i].Length == 8 && currentPath[i].All(char.IsDigit))
                                {
                                    dateFound = currentPath[i];
                                    break;
                                }
                            }
                            
                            string mainFolder = currentPath.Count > 0 ? currentPath[0] : "";
                            
                            viewData[viewName] = new ClashInfo
                            {
                                Title = viewName,
                                MainFolder = mainFolder,
                                IsResolved = isResolved,
                                DateFound = dateFound
                            };

                            string looseName = new string(viewName.Where(char.IsLetterOrDigit).ToArray());
                            if (!string.IsNullOrEmpty(looseName) && !looseViewData.ContainsKey(looseName))
                            {
                                looseViewData[looseName] = viewName;
                            }
                        }
                    }
                }

                var viewpointsNode = doc.Descendants("viewpoints").FirstOrDefault();
                if (viewpointsNode != null)
                {
                    Traverse(viewpointsNode, new List<string>());
                }

                EvaluateStatus(viewData);

                // Parse HTML to get Images
                if (Directory.Exists(htmlFolder))
                {
                    string[] htmlFiles = Directory.GetFiles(htmlFolder, "*.html");
                    foreach (string htmlFile in htmlFiles)
                    {
                        HtmlDocument htmlDoc = new HtmlDocument();
                        htmlDoc.Load(htmlFile, System.Text.Encoding.UTF8);

                        var viewpointNodes = htmlDoc.DocumentNode.SelectNodes("//div[contains(@class, 'viewpoint')]");
                        if (viewpointNodes != null)
                        {
                            foreach (var vp in viewpointNodes)
                            {
                                var h2 = vp.SelectSingleNode(".//h2");
                                string title = h2 != null ? HtmlEntity.DeEntitize(h2.InnerText).Trim() : "";
                                if (!string.IsNullOrEmpty(title))
                                {
                                    title = System.Net.WebUtility.HtmlDecode(title);
                                    title = NormalizeTitle(title);
                                }
                                
                                string looseTitle = new string(title.Where(char.IsLetterOrDigit).ToArray());
                                
                                var img = vp.SelectSingleNode(".//img");
                                string src = img?.GetAttributeValue("src", "");
                                
                                string matchedExactTitle = null;
                                if (viewData.ContainsKey(title)) matchedExactTitle = title;
                                else if (!string.IsNullOrEmpty(looseTitle) && looseViewData.ContainsKey(looseTitle)) matchedExactTitle = looseViewData[looseTitle];

                                if (matchedExactTitle != null)
                                {
                                    if (!string.IsNullOrEmpty(src))
                                    {
                                        viewData[matchedExactTitle].ImagePath = Path.Combine(Path.GetDirectoryName(htmlFile), src);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"XML/HTML Parsing Error:\n{ex.Message}", "Error");
            }

            return viewData;
        }

        public Dictionary<string, ClashInfo> ExtractFromApi(string targetBuilding = null)
        {
            var viewData = new Dictionary<string, ClashInfo>();
            string tempDir = Path.GetTempPath();

            try
            {
                // Method A: Traverse Saved Viewpoints (to match XML logic exactly for Folder/Date info)
                var savedViewpoints = Application.ActiveDocument.SavedViewpoints;
                
                void TraverseNavisworksFolders(SavedItem item, List<string> currentPath)
                {
                    if (item is FolderItem folder)
                    {
                        string folderName = folder.DisplayName.Trim();
                        var newPath = new List<string>(currentPath) { folderName };
                        foreach (var child in folder.Children)
                        {
                            TraverseNavisworksFolders(child, newPath);
                        }
                    }
                    else if (item is SavedViewpoint vp)
                    {
                        string viewName = vp.DisplayName.Trim();
                        bool isResolved = currentPath.Any(f => f.ToUpper() == "RESOLVED");
                        
                        string dateFound = "";
                        for (int i = currentPath.Count - 1; i >= 0; i--)
                        {
                            if (currentPath[i].Length == 8 && currentPath[i].All(char.IsDigit))
                            {
                                dateFound = currentPath[i];
                                break;
                            }
                        }
                        
                        string mainFolder = currentPath.Count > 0 ? currentPath[0] : "";
                        
                        // Export Image to Temp
                        string safeTitle = string.Join("_", viewName.Split(Path.GetInvalidFileNameChars()));
                        string imgPath = Path.Combine(tempDir, safeTitle + ".jpg");
                        
                        try 
                        {
                            // Save current state
                            var activeView = Application.ActiveDocument.ActiveView;
                            
                            // Switch view to the saved viewpoint
                            Application.ActiveDocument.SavedViewpoints.CurrentSavedViewpoint = vp;
                            
                            // Generate image
                            using (System.Drawing.Bitmap bmp = activeView.GenerateImage(ImageGenerationStyle.ScenePlusOverlay, 400, 300))
                            {
                                bmp.Save(imgPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine("Failed to generate image for " + viewName + ": " + ex.Message);
                            imgPath = ""; // If failed, clear path
                        }
                        
                        viewData[viewName] = new ClashInfo
                        {
                            Title = viewName,
                            MainFolder = mainFolder,
                            IsResolved = isResolved,
                            DateFound = dateFound,
                            ImagePath = imgPath
                        };
                    }
                }

                foreach (var item in savedViewpoints.RootItem.Children)
                {
                    if (item is FolderItem folder)
                    {
                        if (!string.IsNullOrEmpty(targetBuilding))
                        {
                            if (folder.DisplayName.Trim() != targetBuilding) continue;
                        }

                         // Filter out default views if necessary
                         TraverseNavisworksFolders(item, new List<string>());
                    }
                }

                // Method B: Also get Clash Detective results as requested
                DocumentClash documentClash = Application.ActiveDocument.GetClash();
                DocumentClashTests clashTests = documentClash.TestsData;
                
                if (clashTests != null && clashTests.Tests != null)
                {
                    foreach (SavedItem testItem in clashTests.Tests)
                    {
                        if (testItem is ClashTest test)
                        {
                            foreach (SavedItem issue in test.Children)
                            {
                                if (issue is ClashResult result)
                                {
                                    string name = result.DisplayName.Trim();
                                    if (!viewData.ContainsKey(name))
                                    {
                                        viewData[name] = new ClashInfo
                                        {
                                            Title = name,
                                            MainFolder = "CLASH CHECK",
                                            IsResolved = (result.Status == ClashResultStatus.Resolved),
                                            DateFound = DateTime.Now.ToString("yyyyMMdd")
                                        };
                                    }
                                }
                            }
                        }
                    }
                }

                EvaluateStatus(viewData);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"API Data Extraction Error:\n{ex.Message}", "Error");
            }

            return viewData;
        }

        private void EvaluateStatus(Dictionary<string, ClashInfo> viewData)
        {
            var validDates = viewData.Values.Select(v => v.DateFound).Where(d => !string.IsNullOrEmpty(d) && d.Length == 8 && d.All(char.IsDigit)).ToList();
            string maxDate = validDates.Count > 0 ? validDates.Max() : "99999999";

            foreach (var kvp in viewData)
            {
                var data = kvp.Value;
                if (data.IsResolved)
                {
                    data.ComputedStatus = "RESOLVED";
                }
                else
                {
                    if (string.IsNullOrEmpty(data.DateFound))
                        data.ComputedStatus = "ACTIVE";
                    else if (data.DateFound == maxDate)
                        data.ComputedStatus = "NEW";
                    else
                        data.ComputedStatus = "ACTIVE";
                }
            }
        }
    }
}
