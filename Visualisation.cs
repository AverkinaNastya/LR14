using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace LR14
{
    public class Visualisation
    {
        public string name;
        public Excel.Worksheet Tablich;
        public void ChtenieInfo(System.Windows.Forms.DataGridView data, System.Windows.Forms.DataVisualization.Charting.Chart StolbInfo)
        {
                Excel.Application ObjWorkExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open("C:/Users/Administrator/Documents/системное программирование/ЛР14/LR14/sheet.xlsx");
            try
            {
                Tablich = (Excel.Worksheet)ObjWorkBook.Sheets[1];
                data.ColumnCount = 10;
                data.RowCount = 10;
                for (int i = 0; i < 10; i++)
                    for (int j = 0; j < 10; j++)
                        data.Rows[i].Cells[j].Value = Tablich.Cells[i + 1, j + 1].Text.ToString();
                StolbInfo.DataSource = data;
            }
            finally
            {
                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
            }
        }
        public void ZapisInfo()
        { 

        }

    }
}