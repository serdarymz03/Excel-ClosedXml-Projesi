using ClosedXML.Excel;
using Entity;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ClosedXml_Project
{
    public partial class FrmWriteExcel : Form
    {
        List<Person> people;
        string fullPath;
        public FrmWriteExcel()
        {
            InitializeComponent();
        }

        private void FrmWriteExcel_Load(object sender, EventArgs e)
        {
             people = new List<Person>()
            {
                 new Person(1,"Melda",new DateTime(1995,10,15)),
                new Person(2, "Kerem",new DateTime(1990,1,25)),
                new Person(3, "Tugay",new DateTime(1993,7,19)),
                new Person(4, "Ümit",new DateTime(1991,5,29)),
                new Person(5, "Damla",new DateTime(1999,9,17)),
                new Person(6, "Gizem",new DateTime(1990,11,23)),
                new Person(7,"Hülya",new DateTime(1989,12,13)),
                new Person(8,"Hüseyin",new DateTime(1988,7,26)),
                new Person(9,"Kemal",new DateTime(2000,1,14)),
                new Person(10,"Selin",new DateTime(1997,10,24))
            };
            DtgPersonList.DataSource = people;
            DtgPersonList.Columns["Id"].HeaderText = "No";
            DtgPersonList.Columns["Name"].HeaderText = "Ad Soyad";
            DtgPersonList.Columns["BirthDate"].HeaderText = "Doğum Tarihi";
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = saveFileDialog1.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {
                fullPath = saveFileDialog1.FileName;

                string oldName = Path.GetFileName(fullPath);
                string newName = Path.GetFileNameWithoutExtension(oldName) + ".xlsx";

                fullPath = fullPath.Replace(oldName, newName);
                TxtPath.Text = fullPath;
            }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (DtgPersonList.DataSource==null)
            {
                return;
            }
            if (string.IsNullOrEmpty(fullPath))
            {
                return;
            }

            IXLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.AddWorksheet("Veriler");

            IXLRow row = worksheet.Row(1);
            row.Cell(1).Value = "No";
            row.Cell(2).Value = "Ad Soyad";
            row.Cell(3).Value = "Doğum Tarihi";

            row = row.RowBelow();

            foreach (DataGridViewRow item in DtgPersonList.Rows)
            {
                row.Cell("A").Value = item.Cells["Id"].Value;
                row.Cell("B").Value = item.Cells["Name"].Value;
                row.Cell("C").Value = item.Cells["BirthDate"].Value;
                row = row.RowBelow();
            }

            /*foreach (Person item in people)
            {
                row.Cell("A").Value = item.Id;
                row.Cell("B").Value = item.Name;
                row.Cell("C").Value = item.BirthDate;
                row = row.RowBelow();
            }*/

            var firstRow = worksheet.FirstRow();

            firstRow.RowUsed().SetAutoFilter(true);
            firstRow.Cells().Style.Font.Bold = true;
            firstRow.Height = firstRow.Height * 2;

            var column = worksheet.Column("C");
            column.Width = 11.5;

            
            worksheet.Protect("12345");           
            workbook.SaveAs(fullPath);
            System.Diagnostics.Process.Start(fullPath);
        }
    }
}
