using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace companyxml
{
    public partial class AL : Form
    {
        public AL()
        {
            InitializeComponent();
        }

        private void AL_Load(object sender, EventArgs e)
        {
            string[] y =
            {         "",
                DateTime.Today.AddYears(1).Year.ToString(),
                DateTime.Today.Year.ToString(),
                DateTime.Today.AddYears(-1).Year.ToString()
            };
            yearC.DataSource = y;
            yearC.SelectedItem = DateTime.Today.Year.ToString();

            monthC.SelectedItem = DateTime.Now.Month.ToString();
            dataGridView1.Font = new Font("標楷體", 17);
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                DataGridViewColumn column = dataGridView1.Columns[i];
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }

            this.dataGridView1.SelectionMode =
            DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.MultiSelect = false;

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            XElement xe = XElement.Load("tsell.xml");
            var n = from m in xe.Elements()
                    where m.Element("Date").Value.Contains(yearC.Text)
                    select new
                    {
                        規格 = m.Element("Size").Value,
                        重量 = Convert.ToDecimal(m.Element("NW").Value),
                        數量 = Convert.ToInt32(m.Element("Quantity").Value)
                    };
            var n1 = n.GroupBy(m => m.規格 ).Select(m => new
            {
                規格 = m.Key,
                重量 = m.Sum(y1 => y1.重量),
                數量 = m.Sum(y1 => y1.數量)
            });
           
                     
                     
            dataGridView1.DataSource = n1.ToList();
        }

        private void monthC_SelectedIndexChanged(object sender, EventArgs e)
        {
            string date = yearC.Text + monthC.Text;
            XElement xe = XElement.Load("tsell.xml");
            var n = from m in xe.Elements()
                    where m.Element("Date").Value.Contains(date)
                    select new
                    {
                        規格 = m.Element("Size").Value,
                        重量 = Convert.ToDecimal(m.Element("NW").Value),
                        數量 = Convert.ToInt32(m.Element("Quantity").Value)
                    };
            var n1 = n.GroupBy(m =>  m.規格 ).Select(m => new
            {
                規格 = m.Key,
                重量 = m.Sum(y1 => y1.重量),
                數量 = m.Sum(y1 => y1.數量)
            });
            dataGridView1.DataSource = n1.ToList();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            monthC.Text = " ";
           
            XElement xe = XElement.Load("tsell.xml");
            var n = from m in xe.Elements()
                    where m.Element("Date").Value.Contains(yearC.Text)
                    select new
                    {
                        規格 = m.Element("Size").Value,
                        重量 = Convert.ToDecimal(m.Element("NW").Value),
                        數量 = Convert.ToInt32(m.Element("Quantity").Value)
                    };
            var n1 = n.GroupBy(m => m.規格 ).Select(m => new
            {
                規格 = m.Key,
                重量 = m.Sum(y1 => y1.重量),
                數量 = m.Sum(y1 => y1.數量)
            });
            dataGridView1.DataSource = n1.ToList();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
