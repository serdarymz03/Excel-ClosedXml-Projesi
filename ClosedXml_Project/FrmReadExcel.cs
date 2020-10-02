using ClosedXML.Excel;
using Entity;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ClosedXml_Project
{
    public partial class FrmReadExcel : Form
    {
        public FrmReadExcel()
        {
            InitializeComponent();
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Dosyası |*.xlsx";
            DialogResult dialogResult = openFileDialog1.ShowDialog();

            if (dialogResult==DialogResult.OK)
            {
                string filePath = openFileDialog1.FileName;
                IXLWorkbook workbook = new XLWorkbook(filePath);
                IXLWorksheet worksheet = workbook.Worksheet("Veriler");

                IXLRange range = worksheet.RangeUsed();

                List<Person> people = new List<Person>();
                int lastRow = range.RowCount();

                foreach (IXLRangeRow item in range.Rows(2,lastRow))
                {                 
                    Person person = new Person(
                        Convert.ToInt32(item.Cell("A").Value),
                        item.Cell("B").Value.ToString(),
                        (DateTime)item.Cell("C").Value
                        );
                    people.Add(person);
                }
                
                DtgPersonList.DataSource = people;
                TxtPath.Text = filePath;
            }
        }
    }
}
