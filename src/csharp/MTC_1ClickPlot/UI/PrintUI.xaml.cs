using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;

namespace MTC_1ClickPlot.UI
{
    public partial class PrintUI : Window
    {
        private Dictionary<string, System.Windows.Controls.CheckBox> checkboxMap = new Dictionary<string, System.Windows.Controls.CheckBox>();
        private Dictionary<string, System.Windows.Controls.CheckBox> sheetCheckboxMap = new Dictionary<string, System.Windows.Controls.CheckBox>();
        
        private List<string> sortedNames = new List<string>();
        private List<Managers.PrintableItem> printableItems = new List<Managers.PrintableItem>();

        public Dictionary<string, string> renameMap = new Dictionary<string, string>(); // To be passed back
        private UIData uiData = new UIData();

        public PrintUI(List<string> printerList)
        {
            InitializeComponent();

            cbPrinter.ItemsSource = printerList;
            if (printerList.Contains(Config.TARGET_PRINTER))
                cbPrinter.SelectedItem = Config.TARGET_PRINTER;
            else if (printerList.Count > 0)
                cbPrinter.SelectedIndex = 0;
        }

        public void InjectRevitData(List<string> printSettings, List<string> viewSets, Dictionary<string, string> catNames, List<Managers.PrintableItem> printableItems)
        {
            sortedNames = catNames.Keys.OrderBy(k => k, new NaturalSortComparer()).ToList();
            this.printableItems = printableItems;

            cbPrintSetting.ItemsSource = printSettings;
            if (printSettings.Count > 0) cbPrintSetting.SelectedIndex = 0;

            cbViewSet.ItemsSource = viewSets;
            if (viewSets.Count > 0) cbViewSet.SelectedIndex = 0;

            RefreshList("");
            RefreshSheetList("");
        }

        private void rbMode_Checked(object sender, RoutedEventArgs e)
        {
            if (panelSetMode == null || panelCustomMode == null) return;
            
            if (rbModeCustom.IsChecked == true)
            {
                panelSetMode.Visibility = Visibility.Collapsed;
                panelCustomMode.Visibility = Visibility.Visible;
            }
            else
            {
                panelSetMode.Visibility = Visibility.Visible;
                panelCustomMode.Visibility = Visibility.Collapsed;
            }
        }

        private void RefreshList(string txt)
        {
            lbCategories.Items.Clear();
            string search = txt.ToLower();
            foreach (string name in sortedNames)
            {
                if (name.ToLower().Contains(search))
                {
                    bool prevChecked = false;
                    if (checkboxMap.ContainsKey(name))
                    {
                        prevChecked = checkboxMap[name].IsChecked ?? false;
                    }
                    else
                    {
                        prevChecked = Config.DEFAULT_EXCLUSIONS.Any(d => name == d || name.Split(new[] { "  >  " }, StringSplitOptions.None).Contains(d));
                    }

                    var cb = new System.Windows.Controls.CheckBox();
                    cb.Content = name;
                    cb.IsChecked = prevChecked;
                    checkboxMap[name] = cb;
                    lbCategories.Items.Add(cb);
                }
            }
        }

        private void txtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshList(txtSearch.Text);
        }

        private void RefreshSheetList(string txt)
        {
            lbSheets.Items.Clear();
            string search = txt.ToLower();
            foreach (var item in printableItems)
            {
                string name = item.Name;
                if (name.ToLower().Contains(search))
                {
                    bool prevChecked = false;
                    if (sheetCheckboxMap.ContainsKey(name))
                    {
                        prevChecked = sheetCheckboxMap[name].IsChecked ?? false;
                    }

                    var cb = new System.Windows.Controls.CheckBox();
                    cb.Content = name;
                    cb.IsChecked = prevChecked;
                    sheetCheckboxMap[name] = cb;
                    lbSheets.Items.Add(cb);
                }
            }
        }

        private void txtSheetSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshSheetList(txtSheetSearch.Text);
        }

        private void btnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var cb in sheetCheckboxMap.Values) cb.IsChecked = true;
        }

        private void btnClearAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var cb in sheetCheckboxMap.Values) cb.IsChecked = false;
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "เลือกโฟลเดอร์สำหรับบันทึกไฟล์ PDF";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtPath.Text = dialog.SelectedPath;
                }
            }
        }

        private void btnRun_Click(object sender, RoutedEventArgs e)
        {
            uiData.ExcludedNames = checkboxMap.Where(kvp => kvp.Value.IsChecked == true).Select(kvp => kvp.Key).ToList();
            
            var selectedSheetNames = sheetCheckboxMap.Where(kvp => kvp.Value.IsChecked == true).Select(kvp => kvp.Key).ToList();
            uiData.SelectedViewIds = printableItems.Where(item => selectedSheetNames.Contains(item.Name)).Select(item => item.Id).ToList();

            uiData.UseCustomMode = rbModeCustom.IsChecked == true;
            uiData.PrinterName = cbPrinter.SelectedItem != null ? cbPrinter.SelectedItem.ToString() : "";
            uiData.PrintSettingName = cbPrintSetting.SelectedItem != null ? cbPrintSetting.SelectedItem.ToString() : "";
            uiData.ViewSetName = cbViewSet.SelectedItem != null ? cbViewSet.SelectedItem.ToString() : "";
            uiData.IsCombine = rbCombine.IsChecked == true;
            uiData.OutPath = txtPath.Text;
            uiData.DoPrint = chkDoPrint.IsChecked == true;
            uiData.AutoRename = chkAutoRename.IsChecked == true;

            this.DialogResult = true;
            this.Close();
        }

        public UIData GetUIData()
        {
            return uiData;
        }
    }

    public class UIData
    {
        public List<string> ExcludedNames { get; set; }
        public List<Autodesk.Revit.DB.ElementId> SelectedViewIds { get; set; }
        public bool UseCustomMode { get; set; }
        public string PrinterName { get; set; }
        public string PrintSettingName { get; set; }
        public string ViewSetName { get; set; }
        public bool IsCombine { get; set; }
        public string OutPath { get; set; }
        public bool DoPrint { get; set; }
        public bool AutoRename { get; set; }

        public UIData()
        {
            ExcludedNames = new List<string>();
            SelectedViewIds = new List<Autodesk.Revit.DB.ElementId>();
        }
    }

    public class NaturalSortComparer : IComparer<string>
    {
        public int Compare(string x, string y)
        {
            return Regex.Replace(x, @"\d+", m => m.Value.PadLeft(10, '0')).CompareTo(Regex.Replace(y, @"\d+", m => m.Value.PadLeft(10, '0')));
        }
    }
}
