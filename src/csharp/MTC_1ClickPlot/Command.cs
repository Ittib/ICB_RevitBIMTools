using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace MTC_1ClickPlot
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                if (doc.IsModifiable)
                {
                    TaskDialog.Show("Error", "Document is currently in a modifiable state. Please finish or cancel the active transaction.");
                    return Result.Failed;
                }

                // 1. Data Prep
                Managers.RevitDataManager revitMgr = new Managers.RevitDataManager(doc);
                List<string> psNames;
                Dictionary<string, ElementId> psIdMap;
                revitMgr.GetPrintSettings(out psNames, out psIdMap);

                List<string> vsNames;
                Dictionary<string, ElementId> vsIdMap;
                revitMgr.GetViewSheetSets(out vsNames, out vsIdMap);

                var printableItems = revitMgr.GetPrintableItems();
                var catNames = revitMgr.GetCategories();
                var printersList = Managers.PrintSystemManager.GetInstalledPrinters();

                // 2. Show UI
                UI.PrintUI ui = new UI.PrintUI(printersList);
                ui.InjectRevitData(psNames, vsNames, catNames, printableItems);

                bool? dialogResult = ui.ShowDialog();
                if (dialogResult == true)
                {
                    string outMessage = "UI Closed";
                    var uiData = ui.GetUIData();
                    System.Collections.Generic.Dictionary<string, string> renameMap = null;

                    // 3. EXECUTION WITH TRANSACTION GROUP
                    using (TransactionGroup tg = new TransactionGroup(doc, "MTC: Smart Print Process"))
                    {
                        tg.Start();
                        try
                        {
                            // 3.1 Change Color
                            using (Transaction tColor = new Transaction(doc, "MTC: Temp Color"))
                            {
                                tColor.Start();
                                revitMgr.ChangeColorsToBlack(uiData.ExcludedNames);
                                tColor.Commit();
                            }

                            // 3.2 Print
                            if (uiData.DoPrint)
                            {
                                using (Transaction tPrint = new Transaction(doc, "MTC: Setup Print"))
                                {
                                    tPrint.Start();
                                    renameMap = revitMgr.ApplyPrintSettingsAndSubmit(uiData, psIdMap, vsIdMap);
                                    tPrint.Commit();
                                }

                                // 3.3 Spooler wait
                                outMessage = Managers.PrintSystemManager.WaitForSpooler(uiData.PrinterName);
                            }
                            else
                            {
                                outMessage = "✅ เปลี่ยนสีเสร็จสิ้น (ไม่ได้สั่ง Print)";
                            }
                        }
                        catch (Exception ex)
                        {
                            outMessage = "❌ Print Error: " + ex.Message;
                        }
                        finally
                        {
                            // 3.4 Rollback (พระเอกของงาน)
                            tg.RollBack();
                        }

                        // 4. Auto-Rename
                        int mapCount = renameMap != null ? renameMap.Count : 0;
                        if (uiData.DoPrint && !uiData.IsCombine && uiData.AutoRename && mapCount > 0)
                        {
                            outMessage += "\n⏳ กำลังรอไฟล์จาก Temp เพื่อเปลี่ยนชื่อและย้ายโฟลเดอร์...";
                            string renameStatus = Managers.PrintSystemManager.WatchAndRenameFiles(renameMap, uiData.OutPath);
                            outMessage += renameStatus;
                        }

                        TaskDialog.Show("MTC 1ClickPlot", outMessage);
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
