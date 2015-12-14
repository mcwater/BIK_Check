using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.dgvData.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                this.button1.Enabled = false;
                System.Data.Odbc.OdbcConnection ocConn =
                    new System.Data.Odbc.OdbcConnection();

                string strTableName = @"bnkseek.dbf";
                string strConn =
                    @" Driver={Microsoft dBASE Driver (*.dbf)}; SourceType=DBF; " +
                    @" Data Source=" + strTableName + "; Exclusive=No; NULL=NO; " +
                    @" Collate=Machine; BACKGROUNDFETCH=NO; DELETE=NO";

                ocConn.ConnectionString = strConn;
                ocConn.Open();

                string strSql = "select NEWNUM AS BIK, NAMEP AS NAME, KSNP AS Acc ,NNP AS City, ADR from bnkseek.dbf";
                //this.textBox1.Text;

                //需要 using System.Data.Odbc;
                OdbcDataAdapter oda = new OdbcDataAdapter(strSql, ocConn);
                DataTable dt = new DataTable();
                oda.Fill(dt);

                dgvData.DataSource = dt;

                ocConn.Close();
                this.button1.Enabled = true;

            }
            catch (Exception ex)
            {
                this.button1.Enabled = true;
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {

                ((Button)sender).Enabled = false;
                string inputBIC = this.textBox1.Text.Trim();
                
                if (string.IsNullOrEmpty(inputBIC))
                {
                    MessageBox.Show("BIC can be empty");
                    ((Button)sender).Enabled = true;
                    return;
                }
                if (!Regex.IsMatch(inputBIC, @"^\d{9}$"))
                {
                    MessageBox.Show("BIC must be 9 digits");
                    ((Button)sender).Enabled = true;

                    return;
                }


                System.Data.Odbc.OdbcConnection ocConn =
                    new System.Data.Odbc.OdbcConnection();

                string strTableName = @"bnkseek.dbf";
                string strConn =
                    @" Driver={Microsoft dBASE Driver (*.dbf)}; SourceType=DBF; " +
                    @" Data Source=" + strTableName + "; Exclusive=No; NULL=NO; " +
                    @" Collate=Machine; BACKGROUNDFETCH=NO; DELETE=NO";

                ocConn.ConnectionString = strConn;
                ocConn.Open();


                string strSql = "select NEWNUM AS BIK, NAMEP AS NAME, KSNP AS Acc ,NNP AS City, ADR from bnkseek.dbf where NEWNUM='" + this.textBox1.Text + "'";


                OdbcDataAdapter oda = new OdbcDataAdapter(strSql, ocConn);
                DataTable dt = new DataTable();
                oda.Fill(dt);

                dgvData.DataSource = dt;

                ocConn.Close();
                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("No record has been found.");
                }
                ((Button)sender).Enabled = true;
            }
            catch (Exception ex)
            {
                ((Button)sender).Enabled = true;
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

            if(this.dgvData.SelectedRows.Count>0)
            {
                String str = "//RU";
                str += dgvData.SelectedRows[0].Cells["BIK"].Value.ToString();
                str = str + "." + dgvData.SelectedRows[0].Cells["Acc"].Value.ToString();
                str = str + Environment.NewLine + dgvData.SelectedRows[0].Cells["NAME"].Value.ToString();
               // str = str + Environment.NewLine + dgvData.SelectedRows[0].Cells["ADR"].Value.ToString() + " ";
                str+=Environment.NewLine+"G."+dgvData.SelectedRows[0].Cells["City"].Value.ToString();
                Clipboard.SetText(str);
                MessageBox.Show("Following Information has been copied:"+Environment.NewLine+str);
            }
            else
            {
                if (this.dgvData.CurrentCell!=null)
                {
                    DataGridViewRow row= this.dgvData.CurrentCell.OwningRow;
                    String str = "//RU";
                    str += row.Cells["BIK"].Value.ToString();
                    str = str + "." + row.Cells["Acc"].Value.ToString();
                    str = str + Environment.NewLine + row.Cells["NAME"].Value.ToString();
                    // str = str + Environment.NewLine + dgvData.SelectedRows[0].Cells["ADR"].Value.ToString() + " ";
                    str += Environment.NewLine + "G." + row.Cells["City"].Value.ToString();
                    Clipboard.SetText(str);
                    MessageBox.Show("Following Information has been copied:" + Environment.NewLine + str);
                }
                else
                {
                    MessageBox.Show("Please select a bank!");
                }
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode==Keys.Enter)
            {
                this.button2_Click(button2, null);
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
        
        private void dgvData_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow row = this.dgvData.CurrentCell.OwningRow;
            String str = "//RU";
            str += row.Cells["BIK"].Value.ToString();
            str = str + "." + row.Cells["Acc"].Value.ToString();
            str = str + Environment.NewLine + row.Cells["NAME"].Value.ToString();
            // str = str + Environment.NewLine + dgvData.SelectedRows[0].Cells["ADR"].Value.ToString() + " ";
            str += Environment.NewLine + "G." + row.Cells["City"].Value.ToString();
            this.textBox2.Text = str;
        }
    }
}
