using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Printing;
using System.Text.RegularExpressions;
using System.Threading;

namespace MTC_1ClickPlot.Managers
{
    public class PrintSystemManager
    {
        public static List<string> GetInstalledPrinters()
        {
            var printers = new List<string>();
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                printers.Add(printer);
            }
            return printers.OrderBy(p => p, new UI.NaturalSortComparer()).ToList();
        }

        public static string WaitForSpooler(string printerName, int timeoutLimit = 120)
        {
            try
            {
                var printServer = new LocalPrintServer();
                var printQueue = printServer.GetPrintQueue(printerName);
                Thread.Sleep(1000);
                DateTime startTime = DateTime.Now;

                while (true)
                {
                    printQueue.Refresh();
                    if (printQueue.NumberOfJobs == 0)
                    {
                        return "✅ พิมพ์เสร็จสมบูรณ์! ระบบกำลังคืนค่าสี...";
                    }
                    if ((DateTime.Now - startTime).TotalSeconds > timeoutLimit)
                    {
                        return "⚠️ Timeout 2 นาที ข้ามการรอและคืนค่าสี...";
                    }
                    Thread.Sleep(1000);
                }
            }
            catch
            {
                Thread.Sleep(5000); // Fallback
                return "✅ ส่งคำสั่ง Plot สำเร็จ (Fallback 5s)";
            }
        }

        public static string WatchAndRenameFiles(Dictionary<string, string> renameMap, string outDir)
        {
            try
            {
                int processed = 0;
                int expected = renameMap.Count;
                int timeout = 600;
                DateTime startTime = DateTime.Now;

                while (processed < expected)
                {
                    if ((DateTime.Now - startTime).TotalSeconds > timeout)
                    {
                        return "\n⚠️ หมดเวลารอไฟล์จาก PDF24 (Timeout)";
                    }

                    if (Directory.Exists(Config.TEMP_PDF_FOLDER))
                    {
                        foreach (string filePath in Directory.GetFiles(Config.TEMP_PDF_FOLDER, "*.pdf"))
                        {
                            string filename = Path.GetFileName(filePath);
                            string sheetNumExtracted = null;

                            foreach (var key in renameMap.Keys)
                            {
                                // ค้นหา Sheet Number ในชื่อไฟล์แบบกว้างๆ (เพราะ Revit อาจตั้งชื่อไฟล์ให้ต่างกัน)
                                if (Regex.IsMatch(filename, string.Format(@"\b{0}\b", Regex.Escape(key))))
                                {
                                    sheetNumExtracted = key;
                                    break;
                                }
                            }

                            if (sheetNumExtracted != null)
                            {
                                string desiredName = renameMap[sheetNumExtracted];
                                string newPath = Path.Combine(outDir, desiredName);

                                if (File.Exists(newPath))
                                {
                                    try { File.Delete(newPath); } catch { }
                                }

                                try
                                {
                                    File.Move(filePath, newPath);
                                    processed++;
                                    renameMap.Remove(sheetNumExtracted);
                                }
                                catch { }
                            }
                        }
                    }
                    Thread.Sleep(1000);
                }

                if (processed > 0)
                {
                    return string.Format("\n✅ เปลี่ยนชื่อและย้ายไฟล์สำเร็จ {0} แผ่น", processed);
                }
                return "";
            }
            catch (Exception ex)
            {
                return string.Format("\n❌ Rename Error: {0}", ex.Message);
            }
        }
    }
}
