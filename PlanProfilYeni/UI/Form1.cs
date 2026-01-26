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
        private Button btnDrawCad;

        private ProgressBar progressBar;
        private Label lblStatus;

        public Form1()
        {
            Text = "PlanProfilYeni - Otomasyon Aracý";
            Width = 800;
            Height = 500;
            StartPosition = FormStartPosition.CenterScreen;

            InitializeUI();
        }

        private void InitializeUI()
        {
            lstExcels = new ListBox { Left = 20, Top = 20, Width = 500, Height = 200, DataSource = _excelFiles };

            btnAddExcel = new Button { Left = 550, Top = 20, Width = 200, Text = "Excel Ekle" };
            btnAddExcel.Click += BtnAddExcel_Click;

            btnRemoveExcel = new Button { Left = 550, Top = 60, Width = 200, Text = "Seçiliyi Sil" };
            btnRemoveExcel.Click += (s, e) => { if (lstExcels.SelectedItem is string path) _excelFiles.Remove(path); };

            btnHydraulicExcel = new Button { Left = 20, Top = 250, Width = 300, Height = 40, Text = "Hidrolik Tablo Excel Oluþtur" };
            btnHydraulicExcel.Click += BtnHydraulicExcel_Click;

            btnDrawCad = new Button { Left = 340, Top = 250, Width = 200, Height = 40, Text = "AutoCAD'e Çiz" };
            btnDrawCad.Click += BtnDrawCad_Click;

            progressBar = new ProgressBar { Left = 20, Top = 320, Width = 730 };
            lblStatus = new Label { Left = 20, Top = 350, Width = 730, Text = "Hazýr" };

            Controls.AddRange(new Control[] { lstExcels, btnAddExcel, btnRemoveExcel, btnHydraulicExcel, btnDrawCad, progressBar, lblStatus });
        }

        private void BtnAddExcel_Click(object sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog { Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx", Multiselect = true };
            if (dlg.ShowDialog() == DialogResult.OK)
                foreach (var file in dlg.FileNames) if (!_excelFiles.Contains(file)) _excelFiles.Add(file);
        }

        private void BtnHydraulicExcel_Click(object sender, EventArgs e)
        {
            if (_excelFiles.Count == 0) { MessageBox.Show("Önce Excel dosyasý seçmelisin."); return; }

            using var saveDlg = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "HidrolikTablo.xlsx" };
            if (saveDlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                progressBar.Style = ProgressBarStyle.Marquee;
                lblStatus.Text = "Excel oluþturuluyor...";

                // HATA DÜZELTME: Application namespace çakýþmasýný önlemek için tam yol
                System.Windows.Forms.Application.DoEvents();

                var useCase = new ExportHydraulicExcelUseCase(new HydraulicExcelService(), new HydraulicReportWriter());

                useCase.Execute(new HydraulicProcessOptions { InputFiles = _excelFiles.ToArray(), OutputExcelPath = saveDlg.FileName });

                lblStatus.Text = "Tamamlandý";
                progressBar.Style = ProgressBarStyle.Blocks;
                MessageBox.Show("Excel baþarýyla oluþturuldu.");
            }
            catch (Exception ex)
            {
                progressBar.Style = ProgressBarStyle.Blocks;
                MessageBox.Show(ex.Message, "Hata");
                lblStatus.Text = "Hata oluþtu";
                // Hatanýn tam detayýný gösteren mesaj kutusu
                string detayliMesaj = $"Hata Mesajý: {ex.Message}\n\n" +
                                      $"Hata Kaynaðý: {ex.Source}\n\n" +
                                      $"Yýðýn (Stack Trace): {ex.StackTrace}";

                MessageBox.Show(detayliMesaj, "Hata Detayý", MessageBoxButtons.OK, MessageBoxIcon.Error);

                // Panoya kopyala ki bana atabilesin
                Clipboard.SetText(detayliMesaj);
            }
        }

        private void BtnDrawCad_Click(object sender, EventArgs e)
        {
            if (_excelFiles.Count == 0) { MessageBox.Show("Lütfen önce listeden Excel dosyasý seçin."); return; }

            try
            {
                lblStatus.Text = "AutoCAD'e baðlanýlýyor ve çiziliyor...";
                progressBar.Style = ProgressBarStyle.Marquee;

                // HATA DÜZELTME: Application namespace çakýþmasýný önlemek için tam yol
                System.Windows.Forms.Application.DoEvents();

                // 22 Sütun sýnýrý (21 veri + 1 bitiþ)
                double[] boundaries = new double[]
                {
                    0.0, 18.0, 28.0, 58.0, 72.0, 84.0, 96.0, 106.0, 119.5, 134.0,
                    144.0, 154.0, 166.0, 181.0, 210.49, 222.49, 232.49, 242.49, 252.49, 266.876,
                    282.279, 295.826
                };

                var cadOptions = new CadHydraulicTableOptions
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
                    TextLayer = "TipKesit-YazýÇizgileri",
                    TextHeight = 2.0,
                    TextAlignmentCenter = AcAlignment.acAlignmentMiddleCenter
                };

                // HATA DÜZELTME: Constructor parametresi eklendi
                var useCase = new PrintHydraulicTableToCadUseCase(new HydraulicExcelService());

                useCase.Execute(_excelFiles.ToArray(), cadOptions);

                lblStatus.Text = "Çizim Tamamlandý";
                progressBar.Style = ProgressBarStyle.Blocks;
                MessageBox.Show("Ýþlem baþarýyla tamamlandý.\nLütfen AutoCAD'i kontrol edin.");
            }
            catch (Exception ex)
            {
                progressBar.Style = ProgressBarStyle.Blocks;
                lblStatus.Text = "Hata oluþtu";
                // Hatanýn tam detayýný gösteren mesaj kutusu
                string detayliMesaj = $"Hata Mesajý: {ex.Message}\n\n" +
                                      $"Hata Kaynaðý: {ex.Source}\n\n" +
                                      $"Yýðýn (Stack Trace): {ex.StackTrace}";

                MessageBox.Show(detayliMesaj, "Hata Detayý", MessageBoxButtons.OK, MessageBoxIcon.Error);

                // Panoya kopyala ki bana atabilesin
                Clipboard.SetText(detayliMesaj);
            }
        }
    }
}