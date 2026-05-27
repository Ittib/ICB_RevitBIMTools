using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Drawing;

namespace NavisworksClashReporter.Logic
{
    public class ExcelReportUpdater
    {
        public void UpdateExcel(string excelPath, string sheetName, Dictionary<string, ClashInfo> clashData)
        {
            string directory = Path.GetDirectoryName(excelPath);
            string filenameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
            string extension = Path.GetExtension(excelPath);
            string savePath = Path.Combine(directory, $"{filenameWithoutExt}_Updated{extension}");

            // Create Backup
            try
            {
                File.Copy(excelPath, savePath, true);
            }
            catch (Exception ex)
            {
                throw new Exception($"ไม่สามารถสร้างไฟล์ Backup ได้ โปรดตรวจสอบว่าไฟล์เป้าหมายถูกเปิดอยู่หรือไม่\n{ex.Message}");
            }

            dynamic app = null;
            dynamic wb = null;
            dynamic ws = null;
            dynamic usedRange = null;

            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) throw new Exception("Excel is not installed on this system.");
                app = Activator.CreateInstance(excelType);
                
                app.Visible = false;
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;

                wb = app.Workbooks.Open(savePath);
                ws = wb.Sheets[sheetName];
                usedRange = ws.UsedRange;

                int maxRow = 1;
                // Find last row with data in column C (Title)
                dynamic lastCellInColC = ws.Range[$"C{ws.Rows.Count}"];
                dynamic lastCell = lastCellInColC.End[-4162]; // Excel.XlDirection.xlUp = -4162
                if (lastCell != null) maxRow = lastCell.Row;
                Marshal.ReleaseComObject(lastCellInColC);

                HashSet<string> existingTitles = new HashSet<string>();

                // 1. Update existing rows
                for (int row = 1; row <= maxRow; row++)
                {
                    dynamic cellC = ws.Range[$"C{row}"];
                    if (cellC.Value2 != null)
                    {
                        string title = cellC.Value2.ToString().Trim();
                        existingTitles.Add(title);

                        if (clashData.ContainsKey(title))
                        {
                            var data = clashData[title];
                            dynamic cellH = ws.Range[$"H{row}"];
                            string currentStatus = cellH.Value2?.ToString()?.Trim();
                            
                            // Skip update if the status in Excel is already "Pending"
                            if (!string.Equals(currentStatus, "Pending", StringComparison.OrdinalIgnoreCase))
                            {
                                if (currentStatus != data.ComputedStatus)
                                {
                                    cellH.Value2 = data.ComputedStatus;
                                }
                            }
                            Marshal.ReleaseComObject(cellH);
                        }
                    }
                    Marshal.ReleaseComObject(cellC);
                }

                // Get Item No
                int itemNo = 1;
                for (int row = maxRow; row > 0; row--)
                {
                    dynamic cellA = ws.Range[$"A{row}"];
                    double parsedNo = 0;
                    if (cellA.Value2 != null && double.TryParse(cellA.Value2.ToString(), out parsedNo))
                    {
                        itemNo = (int)parsedNo + 1;
                        Marshal.ReleaseComObject(cellA);
                        break;
                    }
                    Marshal.ReleaseComObject(cellA);
                }

                // 2. Add New Clashes
                List<ClashInfo> newClashes = new List<ClashInfo>();
                foreach (var kvp in clashData)
                {
                    if (!existingTitles.Contains(kvp.Key))
                    {
                        newClashes.Add(kvp.Value);
                    }
                }

                if (newClashes.Count > 0)
                {
                    int startRow = maxRow + 1;
                    int endRow = maxRow + newClashes.Count;

                    // Copy format from last row
                    if (maxRow > 1)
                    {
                        dynamic sourceRange = ws.Rows[maxRow];
                        dynamic destRange = ws.Range[$"{startRow}:{endRow}"];
                        sourceRange.Copy();
                        destRange.PasteSpecial(-4122); // Excel.XlPasteType.xlPasteFormats = -4122
                        app.CutCopyMode = false;
                        
                        Marshal.ReleaseComObject(sourceRange);
                        Marshal.ReleaseComObject(destRange);
                    }

                    int currentRow = startRow;
                    foreach (var clash in newClashes)
                    {
                        string folderName;
                        if (clash.Title.StartsWith("C")) folderName = "CLASH CHECK";
                        else if (clash.Title.StartsWith("V")) folderName = "VISUAL CHECK";
                        else folderName = "ETC";

                        string dateFound = clash.DateFound;
                        if (!string.IsNullOrEmpty(dateFound) && dateFound.Length == 8)
                        {
                            dateFound = $"{dateFound.Substring(6, 2)}/{dateFound.Substring(4, 2)}/{dateFound.Substring(0, 4)}";
                        }

                        string[] parts = clash.Title.Split('_');
                        string respBy = parts.Length > 1 ? parts[1] : "";
                        string discipline = ""; // Python explicitly uses empty string here
                        string priority = parts.Length > 3 ? parts[3] : "Minor";

                        // A: Item, B: Folder, C: Title, D: Priority, E: Discipline, F: Date, G: Image, H: Status, I: Resp
                        object[,] rowData = new object[1, 9];
                        rowData[0, 0] = itemNo++;
                        rowData[0, 1] = folderName;
                        rowData[0, 2] = clash.Title;
                        rowData[0, 3] = priority;
                        rowData[0, 4] = discipline;
                        rowData[0, 5] = dateFound;
                        rowData[0, 6] = ""; // Image goes here
                        rowData[0, 7] = clash.ComputedStatus;
                        rowData[0, 8] = respBy;

                        dynamic writeRange = ws.Range[$"A{currentRow}:I{currentRow}"];
                        writeRange.Value2 = rowData;
                        Marshal.ReleaseComObject(writeRange);

                        // Insert Image
                        if (!string.IsNullOrEmpty(clash.ImagePath) && File.Exists(clash.ImagePath))
                        {
                            try
                            {
                                using (System.Drawing.Image img = System.Drawing.Image.FromFile(clash.ImagePath))
                                {
                                    double origWidth = img.Width;
                                    double origHeight = img.Height;

                                    double targetHeightPx = 250;
                                    double ratio = origWidth / origHeight;
                                    double targetHeightPt = targetHeightPx * 0.75;
                                    double targetWidthPt = targetHeightPt * ratio;

                                    dynamic rowRange = ws.Rows[currentRow];
                                    rowRange.RowHeight = targetHeightPt + 10;
                                    Marshal.ReleaseComObject(rowRange);

                                    double neededWidthChars = ((targetHeightPx * ratio) - 5) / 7.0;
                                    dynamic colGRange = ws.Columns["G"];
                                    if (neededWidthChars > (double)colGRange.ColumnWidth)
                                    {
                                        colGRange.ColumnWidth = neededWidthChars + 2;
                                    }
                                    Marshal.ReleaseComObject(colGRange);

                                    dynamic cellG = ws.Range[$"G{currentRow}"];
                                    float left = (float)cellG.Left + 5;
                                    float top = (float)cellG.Top + 5;
                                    Marshal.ReleaseComObject(cellG);

                                    dynamic shapesDyn = ws.Shapes;
                                    dynamic picture = shapesDyn.AddPicture(
                                        clash.ImagePath,
                                        0, // msoFalse
                                        -1, // msoCTrue / msoTrue
                                        left, top, (float)targetWidthPt, (float)targetHeightPt);
                                        
                                    picture.Placement = 1; // Excel.XlPlacement.xlMoveAndSize = 1
                                    Marshal.ReleaseComObject(picture);
                                    Marshal.ReleaseComObject(shapesDyn);
                                }
                            }
                            catch
                            {
                                dynamic cellG = ws.Range[$"G{currentRow}"];
                                cellG.Value2 = "(Image Error)";
                                Marshal.ReleaseComObject(cellG);
                            }
                        }
                        else
                        {
                            dynamic rowRange = ws.Rows[currentRow];
                            rowRange.RowHeight = 15;
                            Marshal.ReleaseComObject(rowRange);
                        }

                        currentRow++;
                    }
                }

                wb.Save();
            }
            finally
            {
                // Cleanup COM objects
                try { if (usedRange != null) Marshal.ReleaseComObject(usedRange); } catch { }
                try { if (ws != null) Marshal.ReleaseComObject(ws); } catch { }
                try { 
                    if (wb != null)
                    {
                        wb.Close(false);
                        Marshal.ReleaseComObject(wb);
                    }
                } catch { }
                try { 
                    if (app != null)
                    {
                        app.ScreenUpdating = true;
                        app.Quit();
                        Marshal.ReleaseComObject(app);
                    }
                } catch { }

                // Force Garbage Collection to ensure EXCEL.exe is killed
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
