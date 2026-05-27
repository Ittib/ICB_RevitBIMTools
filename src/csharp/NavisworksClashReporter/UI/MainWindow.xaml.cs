using System;
using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace NavisworksClashReporter.UI
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += MainWindow_Loaded;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            LoadBuildings();
        }

        private void LoadBuildings()
        {
            CmbBuildings.Items.Clear();
            var extractor = new Logic.ClashDataExtractor();
            
            bool isApiMode = RbModeApi.IsChecked == true;
            System.Collections.Generic.List<string> buildings;

            if (isApiMode)
            {
                buildings = extractor.GetBuildingsFromApi();
            }
            else
            {
                string xmlPath = TxtXmlPath.Text;
                if (!string.IsNullOrWhiteSpace(xmlPath) && File.Exists(xmlPath))
                {
                    buildings = extractor.GetBuildingsFromXml(xmlPath);
                }
                else
                {
                    buildings = new System.Collections.Generic.List<string>();
                }
            }

            foreach (var b in buildings)
            {
                CmbBuildings.Items.Add(b);
            }

            if (CmbBuildings.Items.Count > 0)
            {
                CmbBuildings.SelectedIndex = 0;
            }
        }

        private void BtnRefreshBuildings_Click(object sender, RoutedEventArgs e)
        {
            LoadBuildings();
        }

        private void RbMode_Checked(object sender, RoutedEventArgs e)
        {
            if (GrpXmlHtml != null)
            {
                if (RbModeXml.IsChecked == true)
                    GrpXmlHtml.Visibility = Visibility.Visible;
                else
                    GrpXmlHtml.Visibility = Visibility.Collapsed;
                    
                LoadBuildings(); // Reload buildings when mode changes
            }
        }

        private void BtnBrowseExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            dlg.Title = "เลือกไฟล์ Excel Clash Report";

            if (dlg.ShowDialog() == true)
            {
                TxtExcelPath.Text = dlg.FileName;
                TxtExcelPath.Foreground = System.Windows.Media.Brushes.Black;
                LoadExcelSheets(dlg.FileName);
            }
        }

        private void LoadExcelSheets(string filePath)
        {
            dynamic excelApp = null;
            dynamic wb = null;
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) throw new Exception("Excel is not installed.");
                excelApp = Activator.CreateInstance(excelType);
                excelApp.Visible = false;
                wb = excelApp.Workbooks.Open(filePath, ReadOnly: true);

                CmbSheets.Items.Clear();
                foreach (dynamic sheet in wb.Sheets)
                {
                    CmbSheets.Items.Add(sheet.Name);
                }
                if (CmbSheets.Items.Count > 0) CmbSheets.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ไม่สามารถอ่านไฟล์ Excel ได้:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                try { if (wb != null) { wb.Close(false); System.Runtime.InteropServices.Marshal.ReleaseComObject(wb); } } catch { }
                try { if (excelApp != null) { excelApp.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp); } } catch { }
            }
        }

        private void BtnBrowseXml_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
            dlg.Title = "เลือกไฟล์ XML Viewpoints";

            if (dlg.ShowDialog() == true)
            {
                TxtXmlPath.Text = dlg.FileName;
                TxtXmlPath.Foreground = System.Windows.Media.Brushes.Black;
                LoadBuildings(); // Load buildings after selecting XML
            }
        }

        private void BtnBrowseHtml_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new System.Windows.Forms.FolderBrowserDialog())
            {
                dlg.Description = "เลือกโฟลเดอร์ที่เก็บรูปภาพ HTML";
                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TxtHtmlFolder.Text = dlg.SelectedPath;
                    TxtHtmlFolder.Foreground = System.Windows.Media.Brushes.Black;
                }
            }
        }

        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            string excelPath = TxtExcelPath.Text;
            string sheetName = CmbSheets.SelectedItem?.ToString();

            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            {
                MessageBox.Show("กรุณาเลือกไฟล์ Excel ให้ถูกต้อง", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                MessageBox.Show("กรุณาเลือก Sheet เป้าหมาย", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool isApiMode = RbModeApi.IsChecked == true;
            string xmlPath = TxtXmlPath.Text;
            string htmlFolder = TxtHtmlFolder.Text;

            if (!isApiMode)
            {
                if (string.IsNullOrWhiteSpace(xmlPath) || !File.Exists(xmlPath))
                {
                    MessageBox.Show("กรุณาเลือกไฟล์ XML ให้ถูกต้อง", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (string.IsNullOrWhiteSpace(htmlFolder) || !Directory.Exists(htmlFolder))
                {
                    MessageBox.Show("กรุณาเลือกโฟลเดอร์ HTML ให้ถูกต้อง", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            TxtStatus.Text = "กำลังทำงาน...";
            BtnRun.IsEnabled = false;

            try
            {
                var extractor = new Logic.ClashDataExtractor();
                string selectedBuilding = CmbBuildings.SelectedItem?.ToString();
                
                if (string.IsNullOrWhiteSpace(selectedBuilding))
                {
                    MessageBox.Show("กรุณาเลือกระบุอาคาร (Building)", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    BtnRun.IsEnabled = true;
                    TxtStatus.Text = "รอคำสั่ง...";
                    return;
                }

                var clashData = isApiMode ? extractor.ExtractFromApi(selectedBuilding) : extractor.ExtractFromXmlHtml(xmlPath, htmlFolder, selectedBuilding);

                if (clashData == null || clashData.Count == 0)
                {
                    MessageBox.Show("ไม่พบข้อมูล Clash", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                TxtStatus.Text = "กำลังเขียนข้อมูลลง Excel...";
                var updater = new Logic.ExcelReportUpdater();
                updater.UpdateExcel(excelPath, sheetName, clashData);

                TxtStatus.Text = "เสร็จสมบูรณ์!";
                MessageBox.Show("อัปเดตไฟล์ Excel เสร็จสมบูรณ์", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "เกิดข้อผิดพลาด!";
                MessageBox.Show($"เกิดข้อผิดพลาด:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnRun.IsEnabled = true;
            }
        }
    }
}
