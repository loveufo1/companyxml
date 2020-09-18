using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace companyxml
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public int changecheck = 0;
        int page = 0;
        int countPerPage = 10;
        private void Form1_Load(object sender, EventArgs e)
        {
            XElement xe = XElement.Load("tsell.xml");

            var m = from n in xe.Elements()
                    orderby n.Element("Date").Value, n.Element("OrderNo").Value descending
                    select new
                    {
                        訂單代號 = n.Element("OrderNo").Value,
                        Po = n.Element("Po").Value,
                        出貨日期 = n.Element("Date").Value,
                        狀況 = n.Element("Out").Value,
                        數量 = n.Element("Quantity").Value,
                        規格 = n.Element("Size").Value,
                        商品代號 = n.Element("ItemNo").Value,
                        淨重 = n.Element("NW").Value,
                        總重 = n.Element("GW").Value,
                        港口 = n.Element("Harber").Value,
                        港口代號 = n.Element("HarberCode").Value,
                        成立日期 = n.Element("CDate").Value
                    };
            dataGridView1.DataSource = m.ToList();

            this.dataGridView1.SelectionMode =
           DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.MultiSelect = false;

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
             dataGridView1.Font = new Font("標楷體", 15);

            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                DataGridViewColumn column = dataGridView1.Columns[i];
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }


            string[] y =
           {         "",
                DateTime.Today.AddYears(1).Year.ToString(),
                DateTime.Today.Year.ToString(),
                DateTime.Today.AddYears(-1).Year.ToString()
            };
            comboBox1.DataSource = y;
            comboBox1.SelectedItem = DateTime.Today.Year.ToString();

            comboBox2.SelectedItem = DateTime.Now.Month.ToString();

            comboBox3.SelectedItem = "訂單日期";
            label3.Text = "第" + page.ToString() + "頁";
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
//重新整理
        private void button2_Click(object sender, EventArgs e)
        {
            
            XElement xe = XElement.Load("tsell.xml");

            var dat = from n in xe.Elements()
                      orderby n.Element("Date").Value,n.Element("OrderNo").Value descending
                      select new
                    {
                        訂單代號 = n.Element("OrderNo").Value,
                        Po = n.Element("Po").Value,
                        出貨日期 = n.Element("Date").Value,
                        狀況 = n.Element("Out").Value,
                        數量 = n.Element("Quantity").Value,
                        規格 = n.Element("Size").Value,
                        商品代號 = n.Element("ItemNo").Value,
                        淨重 = n.Element("NW").Value,
                        總重 = n.Element("GW").Value,
                        港口 = n.Element("Harber").Value,
                        港口代號 = n.Element("HarberCode").Value,
                        成立日期 = n.Element("CDate").Value
                    };
            int cxp = countPerPage * page;
            dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
            label3.Text = "第" + page.ToString() + "頁";
            page = 0;
            comboBox1.Text = " ";

            comboBox2.Text = " ";
          
            countPerPage = 10;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //年份搜尋
            XElement xe = XElement.Load("tsell.xml");

            var dat = from n in xe.Elements()
                      orderby n.Element("Date").Value, n.Element("OrderNo").Value descending
                      select new
                      {
                          訂單代號 = n.Element("OrderNo").Value,
                          Po = n.Element("Po").Value,
                          出貨日期 = n.Element("Date").Value,
                          狀況 = n.Element("Out").Value,
                          數量 = n.Element("Quantity").Value,
                          規格 = n.Element("Size").Value,
                          商品代號 = n.Element("ItemNo").Value,
                          淨重 = n.Element("NW").Value,
                          總重 = n.Element("GW").Value,
                          港口 = n.Element("Harber").Value,
                          港口代號 = n.Element("HarberCode").Value,
                          成立日期 = n.Element("CDate").Value
                      };
            int cxp = countPerPage * page;
            if (comboBox3.Text == "出貨日期")
            {
                dat = dat.Where(n => n.出貨日期.ToString().Contains(comboBox1.Text));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();

            }
            if (comboBox3.Text == "成立日期")
            {

                dat = dat.Where(n => n.成立日期.ToString().Contains(comboBox1.Text));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();

            }
            label3.Text = "第" + page.ToString() + "頁";
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            XElement xe = XElement.Load("tsell.xml");
            string strID = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();

            //使用LINT從XML檔案中查詢資訊
            var n = (from m in xe.Elements()
                   where m.Element("OrderNo").Value == strID
                   select m).FirstOrDefault();
            changecheck = 1;
            if (n != null)
            {
                comboBox5.Text = n.Element("HarberCode").Value;
                comboBox5.SelectedItem = n.Element("HarberCode").Value;
                textBox1.Text = n.Element("ItemNo").Value;
                textBox8.Text = n.Element("Po").Value;
                textBox7.Text = n.Element("Date").Value;
                textBox9.Text = n.Element("CDate").Value;
                textBox3.Text = n.Element("Quantity").Value;
                textBox2.Text = n.Element("Size").Value;
                textBox4.Text = n.Element("NW").Value;
                textBox5.Text = n.Element("GW").Value;
                textBox6.Text = n.Element("Out").Value;

                textBox11.Text = n.Element("OrderNo").Value;
                textBox10.Text = n.Element("Al").Value;
            }
        }
        //delete
        private void button7_Click(object sender, EventArgs e)
        {
            XElement xe = XElement.Load("tsell.xml");
            string strID = textBox11.Text;
            if(strID != "")
            {
                var n = (from m in xe.Elements()
                where m.Element("OrderNo").Value == strID
                     select m).FirstOrDefault();
                n.Remove();
            xe.Save("tsell.xml");
                MessageBox.Show("刪除成功");
            }
            else
            {
                MessageBox.Show("請輸入值");
            }
            
            
        }

        private void searchB_Click(object sender, EventArgs e)
        {
            XElement xe = XElement.Load("tsell.xml");
            string strID = textBox11.Text;

            //使用LINT從XML檔案中查詢資訊
            var n = (from m in xe.Elements()
                     where m.Element("OrderNo").Value == strID
                     select m).FirstOrDefault();
            changecheck = 1;
            if (n != null)
            {
                comboBox5.Text = n.Element("HarberCode").Value;
                comboBox5.SelectedItem = n.Element("HarberCode").Value;
                textBox1.Text = n.Element("ItemNo").Value;
                textBox8.Text = n.Element("Po").Value;
                textBox7.Text = n.Element("Date").Value;
                textBox9.Text = n.Element("CDate").Value;
                textBox2.Text = n.Element("Quantity").Value;
                textBox3.Text = n.Element("Size").Value;
                textBox4.Text = n.Element("NW").Value;
                textBox5.Text = n.Element("GW").Value;
                textBox6.Text = n.Element("Out").Value;               
                textBox10.Text = n.Element("Al").Value;
            }
        }

        private void updateB_Click(object sender, EventArgs e)
        {
            XElement xe = XElement.Load("tsell.xml");
            string strID = textBox11.Text;
            if (strID != "")
            {
                var n = (from m in xe.Elements()
                         where m.Element("OrderNo").Value == strID
                         select m).FirstOrDefault();
                n.Element("NW").Value =textBox4.Text;
                n.Element("HarberCode").Value=comboBox5.Text;
                n.Element("ItemNo").Value=textBox1.Text;
                n.Element("Po").Value=textBox8.Text;
                n.Element("Date").Value = textBox7.Text;
                n.Element("CDate").Value= textBox9.Text ;
                n.Element("Quantity").Value=textBox2.Text ;
                n.Element("Size").Value=textBox3.Text;              
                n.Element("GW").Value=textBox5.Text;
                n.Element("Out").Value=textBox6.Text;
               n.Element("OrderNo").Value=textBox11.Text;
                n.Element("Al").Value=textBox10.Text;
                xe.Save("tsell.xml");
                MessageBox.Show("修改成功");
                string Filestr = $@"C:\hw4\{textBox8.Text}";

                Excel.Application app = new Excel.Application();
                Excel.Workbook exwb = app.Workbooks.Add();

                Excel.Worksheet ws = new Excel.Worksheet();
                ws = exwb.Worksheets[1];
                ws.Name = "cell";

                if (comboBox5.Text == "66")
                {
                    app.Cells[3, 1] = "66";
                    app.Cells[5, 1] = "LAEM CHABANG";
                    app.Cells[8, 2] = textBox8.Text;
                    app.Cells[19, 1] = "66";
                    app.Cells[21, 1] = "LAEM CHABANG";
                    app.Cells[24, 2] = textBox8.Text;

                }
                else if (comboBox5.Text == "84")
                {
                    app.Cells[3, 1] = "84";
                    app.Cells[5, 1] = "HO CHI MINH";
                    app.Cells[8, 2] = textBox8.Text;
                    app.Cells[19, 1] = "84";
                    app.Cells[21, 1] = "HO CHI MINH";
                    app.Cells[24, 2] = textBox8.Text;

                }
                else if (comboBox5.Text == "SMART TECH")
                {
                    app.Cells[3, 1] = "SMART TECH";
                    app.Cells[5, 1] = "BAVET";
                    app.Cells[8, 2] = textBox8.Text;
                    app.Cells[19, 1] = "SMART TECH";
                    app.Cells[21, 1] = "BAVET";
                    app.Cells[24, 2] = textBox8.Text;
                }


                //右上
                app.Cells[4, 1] = "--------";

                app.Cells[6, 1] = "MADE IN TAIWAN";
                app.Cells[7, 1] = "R.O.C.";
                app.Cells[8, 1] = @"PO:";
                app.Cells[9, 1] = @"C/No.:";
                app.Cells[11, 5] = textBox11.Text;
                app.Cells[11, 1] = "A";

                app.Cells[4, 5] = "ARTICLE:STEEL PIN";
                app.Cells[5, 5] = "ITEM NO. ";
                app.Cells[5, 6] = textBox1.Text;
                app.Cells[6, 5] = "SIZE ";
                app.Cells[6, 6] = textBox2.Text;
                app.Cells[7, 5] = "QNTY ";
                app.Cells[7, 6] = textBox3.Text;
                app.Cells[7, 7] = "PC";
                app.Cells[8, 5] = "N.W";
                app.Cells[8, 6] = textBox4.Text;
                app.Cells[8, 7] = "KG";
                app.Cells[9, 5] = "G.W";
                app.Cells[9, 6] = textBox5.Text;
                app.Cells[9, 7] = " KG";
                app.Cells[3, 5] = textBox7.Text;  //日期


                //右下

                app.Cells[20, 1] = "--------";
                app.Cells[22, 1] = "MADE IN TAIWAN";
                app.Cells[23, 1] = "R.O.C.";
                app.Cells[24, 1] = @"PO:";
                app.Cells[25, 1] = @"C/No.:";
                app.Cells[27, 5] = textBox11.Text;
                app.Cells[27, 1] = "B";

                app.Cells[20, 5] = "ARTICLE:STEEL PIN";
                app.Cells[21, 5] = "ITEM NO. ";
                app.Cells[21, 6] = textBox1.Text;
                app.Cells[22, 5] = "SIZE ";
                app.Cells[22, 6] = textBox2.Text;
                app.Cells[23, 5] = "QNTY ";
                app.Cells[23, 6] = textBox3.Text;
                app.Cells[23, 7] = "PC";
                app.Cells[24, 5] = "N.W";
                app.Cells[24, 6] = textBox4.Text;
                app.Cells[24, 7] = "KG";
                app.Cells[25, 5] = "G.W";
                app.Cells[25, 6] = textBox5.Text;
                app.Cells[25, 7] = "KG";
                app.Cells[19, 5] = textBox7.Text;



                exwb.SaveAs(Filestr);


                ws = null;
                exwb.Close();
                exwb = null;
                app.Quit();
                app = null;

            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //月份搜尋
            XElement xe = XElement.Load("tsell.xml");
            var dat = from n in xe.Elements()
                      orderby n.Element("Date").Value, n.Element("OrderNo").Value descending
                      select new
                      {

                          訂單代號 = n.Element("OrderNo").Value,
                          Po = n.Element("Po").Value,
                          出貨日期 = n.Element("Date").Value,
                          狀況 = n.Element("Out").Value,
                          數量 = n.Element("Quantity").Value,
                          規格 = n.Element("Size").Value,
                          商品代號 = n.Element("ItemNo").Value,
                          淨重 = n.Element("NW").Value,
                          總重 = n.Element("GW").Value,
                          港口 = n.Element("Harber").Value,
                          港口代號 = n.Element("HarberCode").Value,
                          成立日期 = n.Element("CDate").Value
                      };
            string tt = comboBox1.Text + comboBox2.Text;
            int cxp = countPerPage * page;
            if (comboBox3.Text == "訂單日期")
            {
                dat = dat.Where(n => n.出貨日期.ToString().Contains(tt));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();

            }
            if (comboBox3.Text == "成立日期")
            {
                dat = dat.Where(n => n.成立日期.ToString().Contains(tt));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
            }
            label3.Text = "第" + page.ToString() + "頁";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            page += 1;
            XElement xe = XElement.Load("tsell.xml");
            var count = from n in xe.Elements()
                        group n by n.Element("Id").Value
                        into m
                        select new
                        {
                            c = m.Count()
                        };

            int c = Convert.ToInt32(count.Count());
            int cxp = countPerPage * page;

            if (c - (countPerPage * page) < countPerPage)
            {
                countPerPage = c - (countPerPage * page);
                if (countPerPage < 0) countPerPage = 10;
                page--;
            };

            
            var dat = from n in xe.Elements()
                      orderby n.Element("Date").Value, n.Element("OrderNo").Value descending
                      select new
                      {

                          訂單代號 = n.Element("OrderNo").Value,
                          Po = n.Element("Po").Value,
                          出貨日期 = n.Element("Date").Value,
                          狀況 = n.Element("Out").Value,
                          數量 = n.Element("Quantity").Value,
                          規格 = n.Element("Size").Value,
                          商品代號 = n.Element("ItemNo").Value,
                          淨重 = n.Element("NW").Value,
                          總重 = n.Element("GW").Value,
                          港口 = n.Element("Harber").Value,
                          港口代號 = n.Element("HarberCode").Value,
                          成立日期 = n.Element("CDate").Value
                      };

            //抓讀取值
            string tt = comboBox1.Text + comboBox2.Text;
            if (comboBox3.Text == "出貨日期")
            {
                dat = dat.Where(n => n.出貨日期.ToString().Contains(comboBox1.Text));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
            }
            if (comboBox3.Text == "成立日期")
            {
                dat = dat.Where(n => n.成立日期.ToString().Contains(comboBox1.Text));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
            }
            if (comboBox3.Text == "訂單日期")
            {
                dat = dat.Where(n => n.出貨日期.ToString().Contains(tt));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();

            }
            if (comboBox3.Text == "成立日期")
            {
                dat = dat.Where(n => n.成立日期.ToString().Contains(tt));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
            }
            countPerPage = 10;
            label3.Text = "第" + page.ToString() + "頁";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            page -= 1;
            if (page <= 0) page = 0;
            XElement xe = XElement.Load("tsell.xml");
            var dat = from n in xe.Elements()
                      orderby n.Element("Date").Value, n.Element("OrderNo").Value descending
                      select new
                      {
                          訂單代號 = n.Element("OrderNo").Value,
                          Po = n.Element("Po").Value,
                          出貨日期 = n.Element("Date").Value,
                          狀況 = n.Element("Out").Value,
                          數量 = n.Element("Quantity").Value,
                          規格 = n.Element("Size").Value,
                          商品代號 = n.Element("ItemNo").Value,
                          淨重 = n.Element("NW").Value,
                          總重 = n.Element("GW").Value,
                          港口 = n.Element("Harber").Value,
                          港口代號 = n.Element("HarberCode").Value,
                          成立日期 = n.Element("CDate").Value
                      };
            label3.Text = "第" + page.ToString() + "頁";

            //抓讀取值
            string tt = comboBox1.Text + comboBox2.Text;
            int cxp = countPerPage * page;
            if (comboBox3.Text == "出貨日期")
            {
                dat = dat.Where(n => n.出貨日期.ToString().Contains(comboBox1.Text));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
            }
            if (comboBox3.Text == "成立日期")
            {
                dat = dat.Where(n => n.成立日期.ToString().Contains(comboBox1.Text));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
            }
            if (comboBox3.Text == "訂單日期")
            {
                dat = dat.Where(n => n.出貨日期.ToString().Contains(tt));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();

            }
            if (comboBox3.Text == "成立日期")
            {
                dat = dat.Where(n => n.成立日期.ToString().Contains(tt));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
            }
        }

     
        private void button1_Click(object sender, EventArgs e)
        {
            New n = new New();
            n.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Excel.Application app = new Excel.Application();
            string Filestr = $@"C:\hw4\{textBox8.Text}";
            app.Visible = true;

            Excel.Workbook workbook = app.Workbooks.Open(Filestr);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            //app.Workbooks.Open(Filestr);
            
            
            Excel.PageSetup pageSetup = worksheet.PageSetup;
            
            pageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            
            workbook.Save();
            worksheet.PrintOutEx();
            workbook.Close();
            app.Quit();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            AL al = new AL();
            al.Show();
        }
    }
}
