using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TestUpload
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string fileLocation = Server.MapPath("~/App_Data/" + "test.xls");
            string connectionString = @"";
            connectionString = "Provider =Microsoft.Jet.OLEDB.4.0;Data Source=" +
                    fileLocation + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
            //connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
            //          fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            OleDbConnection con = new OleDbConnection(connectionString);

            con.Open();

            con.Close();
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            //string strConnection = "ConnectionString";
            string connectionString = @"Data Source=CO1P8S\DEV_03;Initial Catalog=test;Persist Security Info=True;User ID=adm;Password=adm";
            if (flu_load.HasFile)
            {
                string fileName = Path.GetFileName(flu_load.PostedFile.FileName);
                string fileExtension = Path.GetExtension(flu_load.PostedFile.FileName);
                string fileLocation = Server.MapPath("~/App_Data/" + fileName);
                flu_load.SaveAs(fileLocation);
                if (fileExtension == ".xls")
                {
                    connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                      fileLocation + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                }
                else if (fileExtension == ".xlsx")
                {
                    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                      fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                }

                OleDbConnection con = new OleDbConnection(connectionString);
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandType = CommandType.Text;
                cmd.Connection = con;
                OleDbDataAdapter dAdapter = new OleDbDataAdapter(cmd);
                DataTable dtExcelRecords = new DataTable();
                con.Open();

                DataTable dtExcelSheetName = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string getExcelSheetName = dtExcelSheetName.Rows[0]["Table_Name"].ToString();
                cmd.CommandText = "SELECT * FROM [" + getExcelSheetName + "]";
                dAdapter.SelectCommand = cmd;
                dAdapter.Fill(dtExcelRecords);
                GridView1.DataSource = dtExcelRecords;
                GridView1.DataBind();
            }
        }
    }
}
