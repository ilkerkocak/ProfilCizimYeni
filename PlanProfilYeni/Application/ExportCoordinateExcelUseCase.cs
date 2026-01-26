using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PlanProfilYeni.Domain;
namespace PlanProfilYeni.Application
{
    public class ExportCoordinateExcelUseCase
    {
        public void Execute(List<string> files, HydraulicProcessOptions options)
        {
            MessageBox.Show("ExportCoordinateExcelUseCase çalıştı");
        }
    }
}