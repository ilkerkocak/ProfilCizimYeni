using PlanProfilYeni.Application;
using PlanProfilYeni.Cad;
using PlanProfilYeni.Services;
using System;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using System.Runtime.InteropServices;

namespace PlanProfilYeni.UI
{
    public class Form1 : Form
    {
        private readonly BindingList<string> _excelFiles = new();

        private ListBox lstExcels;
        private Button btnAddExcel;
        private Button btnRemoveExcel;

        private Button btnHydraulicExcel;
        private Button btnHydraulicCad;

        private Button btnCoordinateExcel;
        private Button btnCoordinateCad;

        private Button btnProfileCad;   // ✅ YENİ

        // 🔒 ÖLÇEKLERİ SELECT ETMEK İÇİN
        private ComboBox cmbVerticalScale;
        private ComboBox cmbHorizontalScale;

        private ProgressBar progressBar;
        private Label lblStatus;

        public Form1()
        {
            Text = "PlanProfilYeni - Otomasyon Aracı";
            Width = 900;
            Height = 560;
            StartPosition = FormStartPosition.CenterScreen;

            InitializeUI();
        }

        private void InitializeUI()
        {
            lstExcels = new ListBox
            {
                Left = 20,
                Top = 20,
                Width = 500,
                Height = 200,
                DataSource = _excelFiles
            };

            btnAddExcel = new Button { Left = 550, Top = 20, Width = 300, Text = "Excel Ekle" };
            btnAddExcel.Click += BtnAddExcel_Click;

            btnRemoveExcel = new Button { Left = 550, Top = 60, Width = 300, Text = "Seçiliyi Sil" };
            btnRemoveExcel.Click += (s, e) =>
            {
                if (lstExcels.SelectedItem is string path)
                    _excelFiles.Remove(path);
            };

            // --------------------
            // ÖLÇEK SEÇİMLERİ
            // --------------------
            cmbVerticalScale = new ComboBox
            {
                Left = 20,
                Top = 230,
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Items = { "1/100", "1/500", "1/1000" }
            };
            cmbVerticalScale.SelectedIndex = 0;  // Varsayılan: 1/100

            cmbHorizontalScale = new ComboBox
            {
                Left = 180,
                Top = 230,
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Items = { "1/2000", "1/5000", "1/10000" }
            };
            cmbHorizontalScale.SelectedIndex = 1;  // Varsayılan: 1/5000

            Label lblVertical = new Label { Left = 20, Top = 210, Text = "Düşey Ölçek:" };
            Label lblHorizontal = new Label { Left = 180, Top = 210, Text = "Yatay Ölçek:" };

            // --------------------
            // BUTONLAR
            // --------------------
            btnHydraulicExcel = new Button { Left = 20, Top = 270, Width = 400, Height = 40, Text = "Hidrolik Tablo → Excel" };
            btnHydraulicExcel.Click += BtnHydraulicExcel_Click;

            btnHydraulicCad = new Button { Left = 450, Top = 270, Width = 400, Height = 40, Text = "Hidrolik Tablo → AutoCAD" };
            btnHydraulicCad.Click += BtnHydraulicCad_Click;

            btnCoordinateExcel = new Button { Left = 20, Top = 320, Width = 400, Height = 40, Text = "Koordinat Tablosu → Excel" };
            btnCoordinateExcel.Click += BtnCoordinateExcel_Click;

            btnCoordinateCad = new Button { Left = 450, Top = 320, Width = 400, Height = 40, Text = "Koordinat Tablosu → AutoCAD" };
            btnCoordinateCad.Click += BtnCoordinateCad_Click;

            // ✅ PROFİL ÇİZ BUTONU
            btnProfileCad = new Button
            {
                Left = 20,
                Top = 370,
                Width = 830,
                Height = 40,
                Text = "Profil Çizimi → AutoCAD"
            };
            btnProfileCad.Click += BtnProfileCad_Click;

            progressBar = new ProgressBar { Left = 20, Top = 440, Width = 830 };
            lblStatus = new Label { Left = 20, Top = 470, Width = 830, Text = "Hazır" };

            Controls.AddRange(new Control[]
            {
                lstExcels,
                btnAddExcel,
                btnRemoveExcel,
                lblVertical,
                cmbVerticalScale,
                lblHorizontal,
                cmbHorizontalScale,
                btnHydraulicExcel,
                btnHydraulicCad,
                btnCoordinateExcel,
                btnCoordinateCad,
                btnProfileCad,
                progressBar,
                lblStatus
            });
        }

        // =====================================================
        // EXCEL EKLE
        // =====================================================
        private void BtnAddExcel_Click(object sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog
            {
                Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx",
                Multiselect = true
            };

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                foreach (var file in dlg.FileNames)
                    if (!_excelFiles.Contains(file))
                        _excelFiles.Add(file);
            }
        }

        // =====================================================
        // 1) HİDROLİK → EXCEL
        // =====================================================
        private void BtnHydraulicExcel_Click(object sender, EventArgs e)
        {
            if (_excelFiles.Count == 0) { MessageBox.Show("Excel seçmelisin."); return; }

            using var save = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "HidrolikTablo.xlsx" };
            if (save.ShowDialog() != DialogResult.OK) return;

            RunSafe("Hidrolik tablo Excel oluşturuluyor...", () =>
            {
                var useCase = new ExportHydraulicExcelUseCase(
                    new HydraulicExcelService(),
                    new HydraulicReportWriter());

                useCase.Execute(new HydraulicProcessOptions
                {
                    InputFiles = _excelFiles.ToArray(),
                    OutputExcelPath = save.FileName
                });
            });
        }

        // =====================================================
        // 2) HİDROLİK → AUTOCAD
        // =====================================================
        private void BtnHydraulicCad_Click(object sender, EventArgs e)
        {
            if (_excelFiles.Count == 0) { MessageBox.Show("Excel seçmelisin."); return; }

            RunSafe("Hidrolik tablo AutoCAD'e çiziliyor...", () =>
            {
                double[] boundaries =
                {
                    0.0, 18.0, 28.0, 58.0, 72.0, 84.0, 96.0, 106.0, 119.5, 134.0,
                    144.0, 154.0, 166.0, 181.0, 210.49, 222.49, 232.49, 242.49,
                    252.49, 266.876, 282.279, 295.826
                };

                var options = new CadHydraulicTableOptions
                {
                    BaseX = 0,
                    BaseY = 500,
                    BaseZ = 0,
                    HeaderBlockName = "Hidrolik Basligi",
                    ColumnBoundaries = boundaries,
                    TableWidth = 295.826,
                    RowStep = 3.5,
                    FirstRowOffsetY = 1.75,
                    GridLayer = "Arazi",
                    TextLayer = "TipKesit-YazıÇizgileri",
                    TextHeight = 2.0,
                    TextAlignmentCenter = AcAlignment.acAlignmentMiddleCenter
                };

                var useCase = new PrintHydraulicTableToCadUseCase(new HydraulicExcelService());
                useCase.Execute(_excelFiles.ToArray(), options);
            });
        }

        // =====================================================
        // 3) KOORDİNAT → EXCEL
        // =====================================================
        private void BtnCoordinateExcel_Click(object sender, EventArgs e)
        {
            if (_excelFiles.Count == 0) { MessageBox.Show("Excel seçmelisin."); return; }

            RunSafe("Koordinat tablosu Excel oluşturuluyor...", () =>
            {
                var useCase = new ExportCoordinateExcelUseCase();
                useCase.Execute(_excelFiles.ToArray());
            });
        }

        // =====================================================
        // 4) KOORDİNAT → AUTOCAD
        // =====================================================
        private void BtnCoordinateCad_Click(object sender, EventArgs e)
        {
            if (_excelFiles.Count == 0) { MessageBox.Show("Excel seçmelisin."); return; }

            RunSafe("Koordinat tablosu AutoCAD'e çiziliyor...", () =>
            {
                var useCase = new PrintCoordinateCadUseCase();
                useCase.Execute(_excelFiles.ToArray());
            });
        }

        // =====================================================
        // 5) PROFİL → AUTOCAD  ✅
        // =====================================================
        private void BtnProfileCad_Click(object sender, EventArgs e)
        {
            if (_excelFiles.Count == 0)
            {
                MessageBox.Show("Excel seçmelisin.");
                return;
            }

            RunSafe("Profil AutoCAD'e çiziliyor...", () =>
            {
                var acadApp = (Autodesk.AutoCAD.Interop.AcadApplication)
                    System.Runtime.InteropServices.Marshal
                        .GetActiveObject("AutoCAD.Application");

                var doc = acadApp.ActiveDocument;

                double insertX = 0.0;
                double insertY = 300.0;

                // 🔒 ÖLÇEKLER (DROPDOWN'TAN SEÇILI DEĞERLER)
                string verticalScaleStr = cmbVerticalScale.SelectedItem?.ToString() ?? "1/100";
                string horizontalScaleStr = cmbHorizontalScale.SelectedItem?.ToString() ?? "1/5000";

                // Örnek: "1/100" -> 100, "1/5000" -> 5000
                int dusEyOlcek = int.Parse(verticalScaleStr.Split('/')[1]);
                int yatayOlcek = int.Parse(horizontalScaleStr.Split('/')[1]);

                // 🔒 CAD ÇARPANLARI (METRE CİNSİNDEN)
                // 1000m / scale = CAD birimi per metre
                double cadUnitsPerMeterHorizontal = 1000.0 / yatayOlcek;
                double cadUnitsPerMeterVertical = 1000.0 / dusEyOlcek;

                var useCase = new DrawProfileUseCase();
                foreach (var excelPath in _excelFiles)
                {
                    useCase.Execute(
                        excelPath,
                        doc,
                        insertX,
                        insertY,
                        cadUnitsPerMeterHorizontal,
                        cadUnitsPerMeterVertical,
                        System.IO.Path.GetFileNameWithoutExtension(excelPath));

                    insertY -= 120.0;
                }
            });
        }


        // =====================================================
        // ORTAK TRY / CATCH
        // =====================================================
        private void RunSafe(string status, Action action)
        {
            try
            {
                progressBar.Style = ProgressBarStyle.Marquee;
                lblStatus.Text = status;
                System.Windows.Forms.Application.DoEvents();


                action();

                lblStatus.Text = "Tamamlandı";
                MessageBox.Show("İşlem tamamlandı.");
            }
            catch (Exception ex)
            {
                lblStatus.Text = "Hata oluştu";

                string detay = $"Mesaj: {ex.Message}\n\nStack:\n{ex.StackTrace}";
                MessageBox.Show(detay, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Clipboard.SetText(detay);
            }
            finally
            {
                progressBar.Style = ProgressBarStyle.Blocks;
            }
        }
    }
}
