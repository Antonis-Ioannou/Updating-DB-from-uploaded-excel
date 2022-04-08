using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using DevExpress.XtraEditors;
using DevExpress.XtraSplashScreen;
using System.Data.SqlClient;
using System.Xml.Serialization;
using System.IO;
using System.Globalization;

namespace Updating_DB_from_uploaded_excel
{
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        userSettings userSettings = new userSettings();
        public static string ConnectionString = string.Empty;
        public static string path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Megasoft\CallManagement";

        public Form1()
        {
            InitializeComponent();
        }

        private void GetConnectionString()
        {
            string fullPath = Path.Combine(path, @"appSettings.xml");

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            if (!File.Exists(fullPath))
            {
                File.Create(fullPath);
                return;
            }

            XmlSerializer xmlSerializer = new XmlSerializer(typeof(userSettings));

            using (FileStream fs = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                try
                {
                    userSettings = (userSettings)xmlSerializer.Deserialize(fs);
                    ConnectionString = userSettings.connString;
                }
                catch (Exception e)
                {

                }
            }
        }

        private void btnUpdate(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (!string.IsNullOrEmpty(btnChooseFIle.Text))
            {
                SplashScreenManager.ShowForm(typeof(WaitForm1));
                string query = string.Empty;

                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(btnChooseFIle.Text.ToString());
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                System.Array myvalues;
                myvalues = (System.Array)xlRange.Cells.Value;

                GetConnectionString();
                using (SqlConnection sqlConnection = new SqlConnection(ConnectionString))
                {
                    sqlConnection.Open();
                    string cDay = "";
                    string mDay = "";

                    for (int i = 1; i <= xlRange.Rows.Count; i++)
                    {
                        if (i > 1)
                        {
                            //string creationDay = " ";
                            //string modifiedDay = " ";

                            if (!(myvalues.GetValue(i, 3) == null || Convert.ToString(myvalues.GetValue(i, 3)) == " " || myvalues.GetValue(i, 4) == null || Convert.ToString(myvalues.GetValue(i, 4)) == " " || myvalues.GetValue(i, 5) == null || Convert.ToString(myvalues.GetValue(i, 5)) == " " || myvalues.GetValue(i, 6) == null))
                            {
                                SqlCommand command = new SqlCommand();
                                //---creation day---//
                                //creationDay = (myvalues.GetValue(i, 2).ToString());
                                //cDay = Convert.ToDateTime(creationDay);

                                //---modified day---//
                                //modifiedDay = (myvalues.GetValue(i, 7).ToString());
                                //mDay = Convert.ToDateTime(modifiedDay);

                                if (!toggleForInsertOrUpdate.Checked)
                                {
                                    cDay = DateTime.Now.ToString("G");
                                    //query = "insert into Calls (CreationDate,TypeId,ReceiverId,CallContactId,Notes,ModifiedDate) Values ('" + cDay + "'," + myvalues.GetValue(i, 3) + "," + myvalues.GetValue(i, 4) + "," + myvalues.GetValue(i, 5) + ",'" + myvalues.GetValue(i, 6) + "','" + cDay + "')";
                                    command = new SqlCommand("insert into Calls values(@CreationDate,@TypeId,@ReceiverId,@CallContactId,@Notes,@ModifiedDate)");
                                    command.Parameters.AddWithValue("@CreationDate", DateTime.Parse(cDay));
                                    command.Parameters.AddWithValue("@TypeId", myvalues.GetValue(i, 3));
                                    command.Parameters.AddWithValue("@ReceiverId", myvalues.GetValue(i, 4));
                                    command.Parameters.AddWithValue("@CallContactId", myvalues.GetValue(i, 5));
                                    command.Parameters.AddWithValue("@Notes", myvalues.GetValue(i, 6));
                                    command.Parameters.AddWithValue("@ModifiedDate", DateTime.Parse(cDay));
                                    command.Connection = sqlConnection;
                                    command.ExecuteNonQuery();
                                }
                                else
                                {
                                    mDay = DateTime.Now.ToString("d");
                                    query = "update Calls SET TypeId = " + myvalues.GetValue(i, 3) + ", ReceiverId = " + myvalues.GetValue(i, 4) + ", CallContactId = " + myvalues.GetValue(i, 5) + ", Notes = '" + myvalues.GetValue(i, 6) + "', ModifiedDate = '" + mDay + "' WHERE CallsId = " + myvalues.GetValue(i, 1) + "";
                                }
                                //SqlCommand command = new SqlCommand(query);
                            }
                        }
                    }
                }

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                SplashScreenManager.CloseForm();

                XtraMessageBox.Show("Επιτυχείς ενημέρωση.");
            }
            else
            {
                XtraMessageBox.Show("Επιλέξτε αρχείο");
            }
        }

        private void btnChooseFile_Click(object sender, EventArgs e)
        {
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;
                    btnChooseFIle.Text = filePath;
                }
            }
        }
    }
}
