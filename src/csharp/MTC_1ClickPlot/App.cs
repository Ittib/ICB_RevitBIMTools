using System;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace MTC_1ClickPlot
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            string tabName = "MTC Tools";
            string panelName = "Print Tools";

            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch { } // Tab might already exist

            RibbonPanel panel = null;
            foreach (RibbonPanel p in application.GetRibbonPanels(tabName))
            {
                if (p.Name == panelName)
                {
                    panel = p;
                    break;
                }
            }

            if (panel == null)
            {
                panel = application.CreateRibbonPanel(tabName, panelName);
            }

            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData buttonData = new PushButtonData(
                "cmd1ClickPlot",
                "1Click\nPlot Pro",
                assemblyPath,
                "MTC_1ClickPlot.Command");

            buttonData.ToolTip = "Plot views/sheets with automatic line color conversion and auto-rename.";

            // --- ส่วนที่เพิ่มเข้ามาสำหรับการโหลดรูปภาพ (Icon) ---
            // ข้อควรระวัง: ต้องเปลี่ยน Build Action ของไฟล์รูปให้เป็น "Embedded Resource" ใน Visual Studio ก่อน
            string iconResourceName = "MTC_1ClickPlot.Resources.MTC_1ClickPloticon.png"; // เปลี่ยนชื่อไฟล์ตรงนี้ให้ตรงกับของคุณ
            BitmapImage iconImage = GetEmbeddedImage(iconResourceName);
            if (iconImage != null)
            {
                buttonData.LargeImage = iconImage; // ขนาด 32x32 px สำหรับปุ่มใหญ่
            }

            panel.AddItem(buttonData);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        // ฟังก์ชันช่วยโหลดรูปภาพที่ฝังมากับไฟล์ DLL (Embedded Resource)
        private BitmapImage GetEmbeddedImage(string resourceName)
        {
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (System.IO.Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null) return null;
                    BitmapImage image = new BitmapImage();
                    image.BeginInit();
                    image.StreamSource = stream;
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.EndInit();
                    return image;
                }
            }
            catch
            {
                return null;
            }
        }
    }
}
