﻿using System;
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
                    //XtraMessageBox.Show(e.Message.ToString());
                }
            }
        }

        //private void Form1_Load(object sender, EventArgs e)
        //{
        //    XmlSerializer xmlSerializer = new XmlSerializer(typeof(userSettings));

        //    if (!Directory.Exists(path))
        //        Directory.CreateDirectory(path);

        //    using (FileStream fs = new FileStream(path + "\\appSettings.xml", FileMode.OpenOrCreate, FileAccess.Read, FileShare.Read))
        //    {
        //        userSettings = (userSettings)xmlSerializer.Deserialize(fs);
        //    }
        //}

        private void btnUpdate(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (!string.IsNullOrEmpty(btnChooseFIle.Text))
            {
                SplashScreenManager.ShowForm(typeof(WaitForm1));
                string query = string.Empty;
                DateTime crtnDay = new DateTime();
                int tpId;
                int rcvrId;
                int cllcntctId;
                string nts = string.Empty;
                DateTime mdfDate = new DateTime();

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

                    for (int i = 1; i <= xlRange.Rows.Count; i++)
                    {
                        if (i > 1)
                        {
                            for (int j = 1; j <= xlRange.Columns.Count; j++)
                            {
                                if (j > 1)
                                {
                                    //--- epilogi stilis apo Calls ---//
                                    query = "select CreationDate,TypeId,RecieverId,CallContactId,Notes,ModifiedDate from Calls where CallsId = '" + myvalues.GetValue(i, 1) + "'";
                                    SqlCommand command = new SqlCommand(query);
                                    command.Connection = sqlConnection;
                                    SqlDataReader reader = command.ExecuteReader();
                                    while (reader.Read())
                                    {
                                        crtnDay = (DateTime)reader["CreationDate"];
                                        tpId = (int)reader["TypeId"];
                                        rcvrId = (int)reader["RecieverId"];
                                        cllcntctId = (int)reader["CallContactId"];
                                        nts = reader["Notes"].ToString();
                                        mdfDate = (DateTime)reader["ModifiedDate"];
                                    }
                                    reader.Close();


                                }
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
    }
}
