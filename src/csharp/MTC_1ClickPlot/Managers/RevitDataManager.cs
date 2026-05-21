using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;

namespace MTC_1ClickPlot.Managers
{
    public class PrintableItem
    {
        public string Name { get; set; }
        public ElementId Id { get; set; }
    }

    public class RevitDataManager
    {
        private Document _doc;
        private Dictionary<string, Category> _catObjMap = new Dictionary<string, Category>();

        public RevitDataManager(Document doc)
        {
            _doc = doc;
        }

        public void GetPrintSettings(out List<string> names, out Dictionary<string, ElementId> idMap)
        {
            var psElems = new FilteredElementCollector(_doc).OfClass(typeof(PrintSetting)).ToElements();
            idMap = new Dictionary<string, ElementId>();
            foreach (var ps in psElems)
            {
                if (!string.IsNullOrEmpty(ps.Name)) idMap[ps.Name] = ps.Id;
            }
            names = idMap.Keys.OrderBy(k => k, new UI.NaturalSortComparer()).ToList();
        }

        public void GetViewSheetSets(out List<string> names, out Dictionary<string, ElementId> idMap)
        {
            var vsElems = new FilteredElementCollector(_doc).OfClass(typeof(ViewSheetSet)).ToElements();
            idMap = new Dictionary<string, ElementId>();
            foreach (var vs in vsElems)
            {
                if (!string.IsNullOrEmpty(vs.Name) && !vs.Name.StartsWith("<"))
                    idMap[vs.Name] = vs.Id;
            }
            names = idMap.Keys.OrderBy(k => k, new UI.NaturalSortComparer()).ToList();
        }

        public List<PrintableItem> GetPrintableItems()
        {
            var allSheets = new FilteredElementCollector(_doc).OfClass(typeof(ViewSheet)).ToElements();
            var sheetsData = allSheets.Cast<ViewSheet>().Select(s => new PrintableItem
            {
                Name = string.Format("[Sheet] {0}: {1}", s.SheetNumber, s.Name),
                Id = s.Id
            }).OrderBy(x => x.Name, new UI.NaturalSortComparer()).ToList();

            var allViews = new FilteredElementCollector(_doc).OfClass(typeof(View)).ToElements();
            var viewsData = allViews.Cast<View>()
                .Where(v => !v.IsTemplate && v.CanBePrinted && !(v is ViewSheet))
                .Select(v => new PrintableItem
                {
                    Name = string.Format("[View] {0}", v.Name),
                    Id = v.Id
                }).OrderBy(x => x.Name, new UI.NaturalSortComparer()).ToList();

            sheetsData.AddRange(viewsData);
            return sheetsData;
        }

        public Dictionary<string, string> GetCategories()
        {
            var catNames = new Dictionary<string, string>();
            foreach (Category cat in _doc.Settings.Categories)
            {
                if (cat.Name.ToLower().Contains(".dwg")) continue;
                catNames[cat.Name] = cat.Name;
                _catObjMap[cat.Name] = cat;

                foreach (Category subcat in cat.SubCategories)
                {
                    if (!subcat.Name.ToLower().Contains(".dwg"))
                    {
                        string key = string.Format("{0}  >  {1}", cat.Name, subcat.Name);
                        catNames[key] = key;
                        _catObjMap[key] = subcat;
                    }
                }
            }
            return catNames;
        }

        public void ChangeColorsToBlack(List<string> excludedNames)
        {
            foreach (var kvp in _catObjMap)
            {
                if (!excludedNames.Contains(kvp.Key))
                {
                    try
                    {
                        kvp.Value.LineColor = Config.BLACK_COLOR;
                    }
                    catch { }
                }
            }
            _doc.Regenerate();
        }

        public Dictionary<string, string> ApplyPrintSettingsAndSubmit(UI.UIData uiData, Dictionary<string, ElementId> psIdMap, Dictionary<string, ElementId> vsIdMap)
        {
            PrintManager pm = _doc.PrintManager;
            pm.SelectNewPrintDriver(uiData.PrinterName);
            pm.PrintRange = PrintRange.Select;
            pm.PrintToFile = true;

            string outDir = uiData.OutPath;
            if (!System.IO.Directory.Exists(outDir)) System.IO.Directory.CreateDirectory(outDir);

            pm.CombinedFile = uiData.IsCombine;

            Dictionary<string, string> renameMap = new Dictionary<string, string>();

            if (!pm.CombinedFile && uiData.AutoRename)
            {
                List<ViewSheet> sheetsToPrint = new List<ViewSheet>();
                if (uiData.UseCustomMode)
                {
                    foreach (var vid in uiData.SelectedViewIds)
                    {
                        var v = _doc.GetElement(vid) as ViewSheet;
                        if (v != null) sheetsToPrint.Add(v);
                    }
                }
                else
                {
                    if (vsIdMap.ContainsKey(uiData.ViewSetName))
                    {
                        var vsElem = _doc.GetElement(vsIdMap[uiData.ViewSetName]) as ViewSheetSet;
                        if (vsElem != null)
                        {
                            foreach (View v in vsElem.Views)
                            {
                                if (v is ViewSheet) sheetsToPrint.Add((ViewSheet)v);
                            }
                        }
                    }
                }

                foreach (var sheet in sheetsToPrint)
                {
                    string sheetNum = sheet.SheetNumber;
                    
                    // ใช้ Sheet Number เป็นชื่อไฟล์แทน
                    string cleanVal = Regex.Replace(sheetNum, @"[\\/*?:""<>|]", "");
                    string newFilename = string.Format("{0}.pdf", cleanVal);
                    
                    renameMap[sheetNum] = newFilename;
                }
            }

            if (!pm.CombinedFile)
            {
                if (uiData.AutoRename)
                {
                    if (!System.IO.Directory.Exists(Config.TEMP_PDF_FOLDER))
                        System.IO.Directory.CreateDirectory(Config.TEMP_PDF_FOLDER);
                    pm.PrintToFileName = System.IO.Path.Combine(Config.TEMP_PDF_FOLDER, "Sheet.pdf");
                }
                else
                {
                    pm.PrintToFileName = System.IO.Path.Combine(outDir, "Sheet.pdf");
                }
            }
            else
            {
                pm.PrintToFileName = System.IO.Path.Combine(outDir, "Combined_Print.pdf");
            }

            // Apply Print Setting
            if (psIdMap.ContainsKey(uiData.PrintSettingName))
            {
                var psElem = _doc.GetElement(psIdMap[uiData.PrintSettingName]) as PrintSetting;
                if (psElem != null) pm.PrintSetup.CurrentPrintSetting = psElem;
            }

            // Apply ViewSheetSet
            if (!uiData.UseCustomMode)
            {
                if (vsIdMap.ContainsKey(uiData.ViewSetName))
                {
                    var vsElem = _doc.GetElement(vsIdMap[uiData.ViewSetName]) as ViewSheetSet;
                    if (vsElem != null) pm.ViewSheetSetting.CurrentViewSheetSet = vsElem;
                }
            }
            else
            {
                if (uiData.SelectedViewIds.Count > 0)
                {
                    ViewSheetSetting vss = pm.ViewSheetSetting;
                    ViewSet viewSet = new ViewSet();
                    foreach (var vid in uiData.SelectedViewIds)
                    {
                        var v = _doc.GetElement(vid) as View;
                        if (v != null) viewSet.Insert(v);
                    }
                    vss.CurrentViewSheetSet.Views = viewSet;
                    try { vss.SaveAs("_MTC_Temp_Set"); }
                    catch { vss.Save(); }
                }
                else
                {
                    throw new Exception("⚠️ ไม่ได้เลือก Sheet/View ใดเลยในโหมด Custom");
                }
            }

            pm.SubmitPrint();
            return renameMap;
        }
    }
}
