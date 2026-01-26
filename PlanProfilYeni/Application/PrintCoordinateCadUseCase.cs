using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PlanProfilYeni.Application
{
    public class PrintCoordinateCadUseCase
    {
        public void Execute(List<string> files, HydraulicProcessOptions options)
        {
            MessageBox.Show("PrintCoordinateCadUseCase çalıştı");
        }
    }
}