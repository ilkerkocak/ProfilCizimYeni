using PlanProfilYeni.Application;
using PlanProfilYeni.Cad;
using PlanProfilYeni.Services;
using System;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.Interop.Common;

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

        private ProgressBar progressBar;
        private Label lblStatus;

        public Form1()
        {
            Text = "PlanProfilYeni - Otomasyon Aracı";
            Width = 900;
            Height = 520;
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

            // --- BUTONLAR ---
            btnHydraulicExcel = new Button { Left = 20, Top = 250, Width = 400, Height = 40, Text = "Hidrolik Tablo → Excel" };
            btnHydraulicExcel.Click += BtnHydraulicExcel_Click;

            btnHydraulicCad = new Button { Left = 450, Top = 250, Width = 400, Height = 40, Text = "Hidrolik Tablo → AutoCAD" };
            btnHydraulicCad.Click += BtnHydraulicCad_Click;

            btnCoordinateExcel = new Button { Left = 20, Top = 310, Width = 400, Height = 40, Text = "Koordinat Tablosu → Excel" };
            btnCoordinateExcel.Click += BtnCoordinateExcel_Click;

            btnCoordinateCad = new Button { Left = 450, Top = 310, Width = 400, Height = 40, Text = "Koordinat Tablosu → AutoCAD" };
            btnCoordinateCad.Click += BtnCoordinateCad_Click;

            progressBar = new ProgressBar { Left = 20, Top = 380, Width = 830 };
            lblStatus = new Label { Left = 20, Top = 410, Width = 830, Text = "Hazır" };

            Controls.AddRange(new Control[]
            {
                lstExcels,
                btnAddExcel,
                btnRemoveExcel,
                btnHydraulicExcel,
                btnHydraulicCad,
                btnCoordinateExcel,
                btnCoordinateCad,
                progressBar,
                lblStatus
            });
        }

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

        // ================================
        // 1) HİDROLİK → EXCEL
        // ================================
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

        // ================================
        // 2) HİDROLİK → AUTOCAD
        // ================================
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

        // ================================
        // 3) KOORDİNAT → EXCEL
        // ================================
        private void BtnCoordinateExcel_Click(object sender, EventArgs e)
        {
            if (_excelFiles.Count == 0) { MessageBox.Show("Excel seçmelisin."); return; }

            using var save = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "KoordinatTablosu.xlsx" };
            if (save.ShowDialog() != DialogResult.OK) return;

            RunSafe("Koordinat tablosu Excel oluşturuluyor...", () =>
            {
                var useCase = new ExportCoordinateExcelUseCase();
                useCase.Execute(_excelFiles.ToArray());
            });
        }

        // ================================
        // 4) KOORDİNAT → AUTOCAD
        // ================================
        private void BtnCoordinateCad_Click(object sender, EventArgs e)
        {
            if (_excelFiles.Count == 0) { MessageBox.Show("Excel seçmelisin."); return; }

            RunSafe("Koordinat tablosu AutoCAD'e çiziliyor...", () =>
            {
                var useCase = new PrintCoordinateCadUseCase();
                useCase.Execute(_excelFiles.ToArray());
            });
        }

        // ================================
        // ORTAK TRY/CATCH & UI
        // ================================
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
