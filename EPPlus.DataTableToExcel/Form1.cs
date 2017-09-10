using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EPPlus.DataTableToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var rootFolder = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            //excel path
            string newExcelOutputPath = 
                $"{rootFolder}\\output\\{System.Guid.NewGuid().ToString()}.xlsx";
            //database path
            string mdbPath = $"{rootFolder}\\mdb\\InventSystem.accdb";

          //store the table from mdb
            DataTable dt = new DataTable();

            //read from mdb
            using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mdbPath))
            {
                using (OleDbCommand cmd =
                    new OleDbCommand("SELECT Barcode,ItemName,ItemDescription FROM Items", conn))
                {
                    conn.Open();

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }

            if (dt.Rows.Count > 0)
            {
                //export to Excel
                using (ExcelPackage pck = new ExcelPackage(new FileInfo(newExcelOutputPath)))
                {
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Inventory");
                    ws.Cells["A1"].LoadFromDataTable(dt, true);
                    pck.Save();
                }
            }
            else
            {
                MessageBox.Show("Nothing to export!");
            }
           
        }
    }
}
