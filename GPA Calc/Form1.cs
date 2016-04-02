using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GPA_Calc
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string[] headers = { "学期", "课程", "成绩", "学分", "绩点" };
            foreach (string s in headers)
            {
                DataGridViewColumn c = new DataGridViewTextBoxColumn();
                c.HeaderText = s;
                dataGridView1.Columns.Add(c);
            }
            DataGridViewColumn c2 = new DataGridViewCheckBoxColumn();
            c2.HeaderText = "成绩已出";
            dataGridView1.Columns.Add(c2);
            dataGridView1.Columns[4].ReadOnly = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string data = Clipboard.GetText();
            string[] lines = data.Split(Environment.NewLine.ToCharArray());
            foreach (string line in lines)
            {
                string[] cols = line.Split('\t');
                if (cols.Length < 7 || cols[0] == "学期") continue;
                if (exist(cols[0], cols[2])) continue;
                int index = dataGridView1.Rows.Add();
                DataGridViewRow r = dataGridView1.Rows[index];
                r.HeaderCell.Value = dataGridView1.RowCount.ToString();
                r.Cells[0].Value = cols[0];
                r.Cells[1].Value = cols[2];
                r.Cells[2].Value = cols[4];
                r.Cells[3].Value = cols[6];
                if (!(new string[] { "放弃","缓考","缓修","暂缺"}).Contains(cols[4]))
                    r.Cells[5].Value = true;
            }
        }

        private bool exist(string term, string course)
        {
            foreach (DataGridViewRow r in dataGridView1.Rows)
            {
                if ((string)r.Cells[0].Value == term && (string)r.Cells[1].Value == course) return true;
            }
            return false;
        }

        private void dataGridView1_ColumnSortModeChanged(object sender, DataGridViewColumnEventArgs e)
        {

        }

        private void dataGridView1_Sorted(object sender, EventArgs e)
        {
            foreach (DataGridViewRow r in dataGridView1.Rows)
            {
                r.HeaderCell.Value = (r.Index + 1).ToString();
            }
        }

        private double GetGPA(string grade)
        {
            double g;
            if (Double.TryParse(grade, out g))
            {
                if (g >= 95) grade = "A+";
                else if (g >= 90) grade = "A";
                else if (g >= 85) grade = "A-";
                else if (g >= 82) grade = "B+";
                else if (g >= 78) grade = "B";
                else if (g >= 75) grade = "B-";
                else if (g >= 72) grade = "C+";
                else if (g >= 68) grade = "C";
                else if (g >= 65) grade = "C-";
                else if (g >= 64) grade = "D+";
                else if (g >= 61) grade = "D";
                else if (g >= 60) grade = "D-";
            }
            switch (grade)
            {
                case "通过": return 4.3;
                case "A+": return 4.3;
                case "A": return 4;
                case "A-": return 3.7;
                case "B+": return 3.3;
                case "B": return 3;
                case "B-": return 2.7;
                case "C+": return 2.3;
                case "C": return 2;
                case "C-": return 1.7;
                case "D+": return 1.5;
                case "D": return 1.3;
                case "D-": return 1;
                default: return 0;
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void dataGridView1_Paint(object sender, PaintEventArgs e)
        {
            /*foreach (DataGridViewRow r in dataGridView1.Rows)
            {
                if ((bool)r.Cells[5].EditedFormattedValue)
                {
                    r.DefaultCellStyle.BackColor = Color.White;
                }
                else
                {
                    r.DefaultCellStyle.BackColor = Color.Gray;
                }
            }*/
            ReCalc();
        }

        private void ReCalc()
        {
            double TotalPoint = 0, GPAPoint = 0, GotPoint = 0, Sum = 0,Point,GP;
            foreach (DataGridViewRow r in dataGridView1.Rows)
            {
                GP=GetGPA((string)r.Cells[2].Value);
                r.Cells[4].Value = GP.ToString();
                if((bool)r.Cells[5].EditedFormattedValue)
                {
                    Point = Double.Parse((string)r.Cells[3].Value);
                    TotalPoint += Point;
                    string g = (string)r.Cells[2].Value;
                    if (g != "通过" && g != "不通过")
                    {
                        GPAPoint += Point;
                        Sum += GP * Point;
                        //r.DefaultCellStyle.BackColor = Color.Red;
                    }
                    if (GP > 0)
                    {
                        GotPoint += Point;
                    }
                }
            }
            double GPA = Sum / GPAPoint;
            label1.Text =
                "当前GPA：" + (GPAPoint==0?"-":GPA.ToString()) + Environment.NewLine +
                "已获学分：" + GotPoint.ToString() + " / " + TotalPoint.ToString();

            foreach (DataGridViewRow r in dataGridView1.Rows)
            {
                if ((bool)r.Cells[5].EditedFormattedValue)
                {
                    if ((string)r.Cells[2].Value != "通过" && (string)r.Cells[2].Value != "不通过")
                    {
                        if (Double.Parse((string)r.Cells[4].Value) >= GPA)
                        {
                            r.DefaultCellStyle.BackColor = Color.Red;
                        }
                        else
                        {
                            r.DefaultCellStyle.BackColor = Color.Green;
                        }
                    }
                    else
                    {
                        r.DefaultCellStyle.BackColor = Color.White;
                    }
                }
                else
                {
                    r.DefaultCellStyle.BackColor = Color.Gray;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string data = Clipboard.GetText();
            string[] lines = data.Split(Environment.NewLine.ToCharArray());
            foreach (string line in lines)
            {
                /*string[] cols = line.Split('\t');
                if (cols.Length < 7 || cols[0] == "学期") continue;
                if (exist(cols[0], cols[2])) continue;
                int index = dataGridView1.Rows.Add();
                DataGridViewRow r = dataGridView1.Rows[index];
                r.HeaderCell.Value = dataGridView1.RowCount.ToString();
                r.Cells[0].Value = cols[0];
                r.Cells[1].Value = cols[2];
                r.Cells[2].Value = cols[4];
                r.Cells[3].Value = cols[6];
                if (!(new string[] { "放弃", "缓考", "缓修", "暂缺" }).Contains(cols[4]))
                    r.Cells[5].Value = true;*/
            }
        }
    }
}
