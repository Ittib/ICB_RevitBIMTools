using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using HtmlAgilityPack;

namespace NwdHtmlImportApp
{
    public partial class Form1 : Form
    {
        private TextBox txtExcelPath;
        private ComboBox cmbSheets;
        private TextBox txtHtmlFolder;
        private Button btnBrowseExcel;
        private Button btnBrowseHtml;
        private Button btnRun;
        private Label lblStatus;

        public Form1()
        {
            InitializeComponent();
            SetupUI();
        }

        private void SetupUI()
        {
            this.Text = "Navisworks Clash to Excel (C# Version)";
            this.Size = new Size(650, 420);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            int padding = 20;
            int yPos = 20;

            Label lblStep1 = new Label() { Text = "1. เลือกไฟล์ Excel Report:", Font = new Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(padding, yPos), AutoSize = true };
            this.Controls.Add(lblStep1);
            
            btnBrowseExcel = new Button() { Text = "Browse Excel", Location = new Point(250, yPos - 5), Size = new Size(120, 30) };
            btnBrowseExcel.Click += BtnBrowseExcel_Click;
            this.Controls.Add(btnBrowseExcel);

            yPos += 35;
            txtExcelPath = new TextBox() { Location = new Point(padding, yPos), Size = new Size(580, 25), ReadOnly = true };
            this.Controls.Add(txtExcelPath);

            yPos += 40;

            Label lblStep2 = new Label() { Text = "2. เลือก Sheet เป้าหมาย:", Font = new Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(padding, yPos), AutoSize = true };
            this.Controls.Add(lblStep2);

            cmbSheets = new ComboBox() { Location = new Point(250, yPos - 2), Size = new Size(250, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            this.Controls.Add(cmbSheets);

            yPos += 50;

            Label lblStep3 = new Label() { Text = "3. เลือกโฟลเดอร์ HTML:", Font = new Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(padding, yPos), AutoSize = true };
            this.Controls.Add(lblStep3);

            btnBrowseHtml = new Button() { Text = "Browse Folder", Location = new Point(250, yPos - 5), Size = new Size(120, 30) };
            btnBrowseHtml.Click += BtnBrowseHtml_Click;
            this.Controls.Add(btnBrowseHtml);

            yPos += 35;
            txtHtmlFolder = new TextBox() { Location = new Point(padding, yPos), Size = new Size(580, 25), ReadOnly = true };
            this.Controls.Add(txtHtmlFolder);

            yPos += 50;

            btnRun = new Button() { 
                Text = "อัปเดตข้อมูลลง Excel", 
                Location = new Point(padding, yPos), 
                Size = new Size(580, 45),
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                BackColor = Color.LightGreen
            };
            btnRun.Click += BtnRun_Click;
            this.Controls.Add(btnRun);

            yPos += 55;
            lblStatus = new Label() { Text = "พร้อมใช้งาน", Location = new Point(padding, yPos), AutoSize = true, ForeColor = Color.Gray };
            this.Controls.Add(lblStatus);
        }

        private void BtnBrowseExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                ofd.Title = "เลือกไฟล์ Excel Clash Report";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtExcelPath.Text = ofd.FileName;
                    LoadExcelSheets(ofd.FileName);
                }
            }
        }

        private void LoadExcelSheets(string filePath)
        {
            cmbSheets.Items.Clear();
            dynamic app = null;
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    MessageBox.Show("ไม่พบโปรแกรม Excel ในเครื่องนี้", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                app = Activator.CreateInstance(excelType);
                app.Visible = false;
                
                dynamic wb = app.Workbooks.Open(filePath, ReadOnly: true);
                foreach (dynamic sheet in wb.Sheets)
                {
                    cmbSheets.Items.Add(sheet.Name);
                }
                
                if (cmbSheets.Items.Count > 0) cmbSheets.SelectedIndex = 0;
                wb.Close(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ไม่สามารถอ่านชีตจาก Excel ได้: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                }
            }
        }

        private void BtnBrowseHtml_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "เลือกโฟลเดอร์ที่เก็บไฟล์ HTML";
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    txtHtmlFolder.Text = fbd.SelectedPath;
                }
            }
        }

        private void BtnRun_Click(object sender, EventArgs e)
        {
            string excelPath = txtExcelPath.Text;
            string targetSheet = cmbSheets.SelectedItem?.ToString();
            string htmlFolder = txtHtmlFolder.Text;

            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
            {
                MessageBox.Show("กรุณาเลือกไฟล์ Excel ที่ถูกต้อง", "คำเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrEmpty(targetSheet))
            {
                MessageBox.Show("กรุณาเลือก Sheet", "คำเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrEmpty(htmlFolder) || !Directory.Exists(htmlFolder))
            {
                MessageBox.Show("กรุณาเลือกโฟลเดอร์ HTML", "คำเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnRun.Enabled = false;
            lblStatus.Text = "กำลังทำงาน... โปรดรอสักครู่";
            Application.DoEvents();

            try
            {
                ProcessFiles(excelPath, targetSheet, htmlFolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show("เกิดข้อผิดพลาด: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnRun.Enabled = true;
                lblStatus.Text = "เสร็จสมบูรณ์";
            }
        }

        private void ProcessFiles(string excelPath, string targetSheet, string htmlFolder)
        {
            string[] htmlFiles = Directory.GetFiles(htmlFolder, "*.html");
            if (htmlFiles.Length == 0)
            {
                MessageBox.Show("ไม่พบไฟล์ .html ในโฟลเดอร์", "ข้อมูล", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var newClashes = new List<(string Title, string ImagePath)>();

            foreach (var htmlFile in htmlFiles)
            {
                string dir = Path.GetDirectoryName(htmlFile);
                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.Load(htmlFile, System.Text.Encoding.UTF8);

                var viewpoints = doc.DocumentNode.SelectNodes("//div[contains(@class, 'viewpoint')]");
                if (viewpoints != null)
                {
                    foreach (var vp in viewpoints)
                    {
                        var titleNode = vp.SelectSingleNode(".//h2");
                        string title = titleNode != null ? titleNode.InnerText.Trim() : "";

                        var imgNode = vp.SelectSingleNode(".//img");
                        string imgSrc = imgNode != null ? imgNode.GetAttributeValue("src", "") : "";

                        string fullImgPath = "";
                        if (!string.IsNullOrEmpty(imgSrc))
                        {
                            fullImgPath = Path.Combine(dir, imgSrc);
                        }

                        if (!string.IsNullOrEmpty(title))
                        {
                            newClashes.Add((title, fullImgPath));
                        }
                    }
                }
            }

            string ext = Path.GetExtension(excelPath);
            string baseName = excelPath.Substring(0, excelPath.Length - ext.Length);
            string savePath = $"{baseName}_Updated{ext}";
            File.Copy(excelPath, savePath, true);

            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            dynamic app = Activator.CreateInstance(excelType);
            app.Visible = false;
            app.ScreenUpdating = false;

            try
            {
                dynamic wb = app.Workbooks.Open(savePath);
                dynamic ws = wb.Sheets[targetSheet];

                int xlUp = -4162;
                int maxRow = ws.Range["C" + ws.Rows.Count].End(xlUp).Row;

                HashSet<string> existingTitles = new HashSet<string>();
                for (int r = 1; r <= maxRow; r++)
                {
                    var val = ws.Range["C" + r].Value;
                    if (val != null)
                    {
                        existingTitles.Add(val.ToString().Trim());
                    }
                }

                int itemNo = 1;
                for (int r = maxRow; r >= 1; r--)
                {
                    var val = ws.Range["A" + r].Value;
                    string valStr = Convert.ToString(val);
                    if (!string.IsNullOrEmpty(valStr) && double.TryParse(valStr, out double num))
                    {
                        itemNo = (int)num + 1;
                        break;
                    }
                }

                var clashesToAdd = newClashes.Where(c => !existingTitles.Contains(c.Title)).ToList();
                int addedCount = clashesToAdd.Count;

                if (addedCount > 0)
                {
                    int startRow = maxRow + 1;
                    int endRow = maxRow + addedCount;

                    object[,] batchData = new object[addedCount, 8];
                    
                    for (int i = 0; i < addedCount; i++)
                    {
                        var clash = clashesToAdd[i];
                        string folderName = "ETC.";
                        if (clash.Title.StartsWith("C")) folderName = "CLASH CHECK";
                        else if (clash.Title.StartsWith("V")) folderName = "VISUAL CHECK";

                        batchData[i, 0] = itemNo++;
                        batchData[i, 1] = folderName;
                        batchData[i, 2] = clash.Title;
                        batchData[i, 7] = "NEW"; // Column H is index 7
                    }

                    if (maxRow > 1)
                    {
                        ws.Range[$"{maxRow}:{maxRow}"].Copy();
                        int xlPasteFormats = -4122;
                        ws.Range[$"{startRow}:{endRow}"].PasteSpecial(xlPasteFormats);
                    }

                    ws.Range[$"A{startRow}:H{endRow}"].Value = batchData;

                    for (int i = 0; i < addedCount; i++)
                    {
                        int currentRow = startRow + i;
                        var clash = clashesToAdd[i];

                        if (!string.IsNullOrEmpty(clash.ImagePath) && File.Exists(clash.ImagePath))
                        {
                            try
                            {
                                string absPath = Path.GetFullPath(clash.ImagePath);
                                using (var img = System.Drawing.Image.FromFile(absPath))
                                {
                                    float origWidth = img.Width;
                                    float origHeight = img.Height;

                                    float targetHeightPx = 250;
                                    float ratio = origWidth / origHeight;
                                    float targetHeightPt = targetHeightPx * 0.75f;
                                    float targetWidthPt = targetHeightPt * ratio;

                                    ws.Range[$"{currentRow}:{currentRow}"].RowHeight = targetHeightPt + 10;

                                    float neededWidth = ((targetHeightPx * ratio) - 5) / 7f;
                                    float currentWidth = (float)ws.Range["G:G"].ColumnWidth;
                                    if (neededWidth > currentWidth)
                                    {
                                        ws.Range["G:G"].ColumnWidth = neededWidth + 2;
                                    }

                                    dynamic cellG = ws.Range[$"G{currentRow}"];
                                    float left = (float)cellG.Left + 5;
                                    float top = (float)cellG.Top + 5;

                                    dynamic shapes = ws.Shapes;
                                    dynamic pic = shapes.AddPicture(
                                        absPath,
                                        0, // msoFalse
                                        -1, // msoTrue
                                        left,
                                        top,
                                        targetWidthPt,
                                        targetHeightPt
                                    );
                                    
                                    // 1 = xlMoveAndSize
                                    pic.Placement = 1;
                                }
                            }
                            catch (Exception imgErr)
                            {
                                string filename = Path.GetFileName(clash.ImagePath);
                                ws.Range[$"G{currentRow}"].Value = $"{filename} (Err: {imgErr.Message})";
                                ws.Range[$"{currentRow}:{currentRow}"].RowHeight = 15;
                            }
                        }
                        else
                        {
                            ws.Range[$"{currentRow}:{currentRow}"].RowHeight = 15;
                        }
                    }
                }

                wb.Save();
                wb.Close();

                MessageBox.Show($"อัปเดตเรียบร้อย!\n\nเพิ่ม Clash ใหม่จำนวน {addedCount} รายการ\nไฟล์ถูกบันทึกที่:\n{savePath}", "สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                app.ScreenUpdating = true;
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            }
        }
    }
}
