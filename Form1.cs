using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace excel_extractor
{
    public partial class Form1 : Form
    {
        string sConnectionString = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|Excel File < 2003 (*.xls)|*.xls";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    var arr = filePath.Split('.');
                    if (arr.Length > 0)
                    {
                        if (arr[arr.Length - 1] == "xls")
                        {

                            sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                            filePath + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'";
                        }
                        else if (arr[arr.Length - 1] == "xlsx")
                        {
                            sConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0 Xml;HDR=YES';";
                        }
                        FillData();
                    }
                }
            }
        }

        private void FillData()
        {
            if (sConnectionString.Length > 0)
            {
                OleDbConnection cn = new OleDbConnection(sConnectionString);
              try   {
                    cn.Open();
                    DataTable dt = new DataTable();
                    DataTable dbSchema = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (dbSchema == null || dbSchema.Rows.Count < 1)
                    {
                        throw new Exception("Error: Could not determine the name of the first worksheet.");
                    }
                    string firstSheetName = dbSchema.Rows[0]["TABLE_NAME"].ToString();
                    OleDbDataAdapter Adpt = new OleDbDataAdapter("select * from [Daily Cost$]", cn);
                    Adpt.Fill(dt);
                    this.dataGridView1.DataSource = dt;
                }
         catch (Exception ex)
                {
                    Console.WriteLine(ex);
                };
            }
        }
    }
}
