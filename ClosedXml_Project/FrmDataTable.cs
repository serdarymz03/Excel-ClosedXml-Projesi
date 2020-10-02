using Business;
using ClosedXML.Excel;
using DataAccess;
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
    public partial class FrmDataTable : Form
    {
        DataTable peopleTable;
        string fullPath;
        public FrmDataTable()
        {
            InitializeComponent();
            peopleTable = PersonManager.GetInstance().GetPeopleAsTable();
            DtgPersonList.DataSource = peopleTable;
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
            if (DtgPersonList.DataSource == null)
            {
                return;
            }
            if (string.IsNullOrEmpty(fullPath))
            {
                return;
            }

            peopleTable.Columns["ID"].ColumnName = "No";
            peopleTable.Columns["NAME"].ColumnName = "Ad Soyad";
            peopleTable.Columns["BIRTHDATE"].ColumnName = "Doğum TARİHİ";

            IXLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.AddWorksheet(peopleTable, "Veriler");

            workbook.SaveAs(fullPath);
            System.Diagnostics.Process.Start(fullPath);
        }
    }
}
