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

        public static DataTable ImportToDataTable(string SheetName)
        {
            var rootFolder = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            DataTable dt = new DataTable();
            FileInfo fi = new FileInfo($"{rootFolder}\\mdb\\test.xlsx");

            // Check if the file exists
            if (!fi.Exists)
                throw new Exception("File Does Not Exists");

            using (ExcelPackage xlPackage = new ExcelPackage(fi))
            {
                // get the first worksheet in the workbook
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[SheetName];

                // Fetch the WorkSheet size
                ExcelCellAddress startCell = worksheet.Dimension.Start;
                ExcelCellAddress endCell = worksheet.Dimension.End;

                // create all the needed DataColumn
                for (int col = startCell.Column; col <= endCell.Column; col++)
                    dt.Columns.Add(col.ToString());

                // place all the data into DataTable
                for (int row = startCell.Row; row <= endCell.Row; row++)
                {
                    DataRow dr = dt.NewRow();
                    int x = 0;
                    for (int col = startCell.Column; col <= endCell.Column; col++)
                    {
                        dr[x++] = worksheet.Cells[row, col].Value;
                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
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

        private void button2_Click(object sender, EventArgs e)
        {
            ImportToDataTable("Sheet2");
        }
    }
}
