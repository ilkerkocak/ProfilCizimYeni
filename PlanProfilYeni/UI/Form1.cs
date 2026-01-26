using PlanProfilYeni.Application;
using PlanProfilYeni.Services;
using System;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;

namespace PlanProfilYeni.UI
{
    public class Form1 : Form
    {
        private readonly BindingList<string> _excelFiles = new();

        private ListBox lstExcels;
        private Button btnAddExcel;
        private Button btnRemoveExcel;
        private Button btnHydraulicExcel;

        private ProgressBar progressBar;
        private Label lblStatus;

        public Form1()
        {
            Text = "PlanProfilYeni";
            Width = 800;
            Height = 500;

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

            btnAddExcel = new Button
            {
                Left = 550,
                Top = 20,
                Width = 200,
                Text = "Excel Ekle"
            };
            btnAddExcel.Click += BtnAddExcel_Click;

            btnRemoveExcel = new Button
            {
                Left = 550,
                Top = 60,
                Width = 200,
                Text = "Seçiliyi Sil"
            };
            btnRemoveExcel.Click += (s, e) =>
            {
                if (lstExcels.SelectedItem is string path)
                    _excelFiles.Remove(path);
            };

            btnHydraulicExcel = new Button
            {
                Left = 20,
                Top = 250,
                Width = 300,
                Height = 40,
                Text = "Hidrolik Tablo Excel Oluþtur"
            };
            btnHydraulicExcel.Click += BtnHydraulicExcel_Click;

            progressBar = new ProgressBar
            {
                Left = 20,
                Top = 320,
                Width = 730
            };

            lblStatus = new Label
            {
                Left = 20,
                Top = 350,
                Width = 730,
                Text = "Hazýr"
            };

            Controls.AddRange(new Control[]
            {
                lstExcels,
                btnAddExcel,
                btnRemoveExcel,
                btnHydraulicExcel,
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

            if (dlg.ShowDialog() != DialogResult.OK)
                return;

            foreach (var file in dlg.FileNames)
                if (!_excelFiles.Contains(file))
                    _excelFiles.Add(file);
        }

        private void BtnHydraulicExcel_Click(object sender, EventArgs e)
        {
            if (_excelFiles.Count == 0)
            {
                MessageBox.Show("Önce Excel dosyasý seçmelisin.");
                return;
            }

            using var saveDlg = new SaveFileDialog
            {
                Filter = "Excel (*.xlsx)|*.xlsx",
                FileName = "HidrolikTablo.xlsx"
            };

            if (saveDlg.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                progressBar.Style = ProgressBarStyle.Marquee;
                lblStatus.Text = "Excel oluþturuluyor...";

                var useCase = new ExportHydraulicExcelUseCase(
                    new HydraulicExcelService(),
                    new HydraulicReportWriter()
                );

                useCase.Execute(new HydraulicProcessOptions
                {
                    InputFiles = _excelFiles.ToArray(),
                    OutputExcelPath = saveDlg.FileName
                });

                lblStatus.Text = "Tamamlandý";
                MessageBox.Show("Excel baþarýyla oluþturuldu.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Hata");
                lblStatus.Text = "Hata oluþtu";
            }
            finally
            {
                progressBar.Style = ProgressBarStyle.Blocks;
            }
        }
    }
}
