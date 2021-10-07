using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LR14
{
    public class Stolb : Visualisation
    {
        public void StolbInform(System.Windows.Forms.DataVisualization.Charting.Chart StolbInfo, System.Windows.Forms.DataGridView Date)
        {
            StolbInfo.DataSource = Date;
        }
    }
}