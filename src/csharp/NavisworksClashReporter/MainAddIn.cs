using System;
using Autodesk.Navisworks.Api.Plugins;

namespace NavisworksClashReporter
{
    [PluginAttribute("NavisworksClashReporter", "MTC", DisplayName = "Export Clash\nto Excel", ToolTip = "Export Clash Data to Excel directly or via XML/HTML")]
    [AddInPluginAttribute(AddInLocation.AddIn)]
    public class MainAddIn : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            try
            {
                var mainWindow = new UI.MainWindow();
                // ShowDialog to keep Navisworks locked while interacting
                mainWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"เกิดข้อผิดพลาดในการเปิด Add-in:\n{ex.Message}", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            
            return 0;
        }
    }
}
