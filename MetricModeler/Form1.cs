using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace MetricModeler
{
    public partial class Form1 : Form
    {
        List<ProjectHistory> projectHistoryList;
        List<Language> languageList;
        public Form1()
        {
            InitializeComponent();
            projectHistoryList = new List<ProjectHistory>();
            languageList = new List<Language>();
            // Read project history data
            Console.WriteLine("------------ Project History data ----------------");
            readProjectHistory();
            // Read language prod data
            Console.WriteLine("\n------------ Language data ----------------");
            readLanguageList();
        }


        private void readProjectHistory() 
        {
            // Connection string
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\".\\resource\\ProjectHistory.accdb\"";

            // Create a connection
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                string strSQL = "SELECT * FROM ProjectHistory";
                OleDbCommand command = new OleDbCommand(strSQL, connection);
                try
                {
                    // Open connecton    
                    connection.Open();
                    // Execute command    
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            projectHistoryList.Add(new ProjectHistory(
                                reader["Project #"].ToString(), 
                                reader["Project Name"].ToString(),
                                reader["Project Description"].ToString(),
                                reader["Project Type"].ToString(),
                                (DateTime)reader["Start Date"],
                                (DateTime)reader["End Date"],
                                (int)reader["Est Duration"],
                                (int)reader["Est Project Cost"],
                                (int)reader["Actual Project Cost"],
                                (int)reader["Est Effort"], 
                                (int)reader["Actual Effort"],
                                (int)reader["Est LOC"],
                                (int)reader["Actual LOC"],
                                (int)reader["Estimated FP"],
                                (int)reader["Actual FP"],
                                (int)reader["Expected Error Rate"],
                                (int)reader["Ave Cost per Person Hour"],
                                (int)reader["Average Staffing Level"],
                                (int)reader["Design Review Hours"],
                                (int)reader["Errors Found"],
                                (int)reader["Defects Reported"],
                                reader["Development Language"].ToString(),
                                (int)reader["Language Productivity Factor"],
                                (int)reader["CPM Tasks Defined"],
                                (int)reader["Change Orders Issued"],
                                (int)reader["Documentation Pages"]
                                )
                            );
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            foreach (ProjectHistory p in projectHistoryList)
                Console.WriteLine(p.ToString());
        }

        private void readLanguageList() 
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var executablePath = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(System.IO.Path.Combine(executablePath, "resource", "language_prod.xlsx"))))
            {
                var languageSheet = xlPackage.Workbook.Worksheets.First();
                var totalRows = languageSheet.Dimension.End.Row; // retrieve the last row
                var totalColumns = languageSheet.Dimension.End.Column; // retrieve the last col

                // start from B4
                for (int rowNum = 4; rowNum <= totalRows; rowNum++)
                {
                    var row = languageSheet.Cells[rowNum, 2, rowNum, totalColumns].Select(c => c.Value);
                    // read all the language into list
                    languageList.Add(new Language(row.ToArray()[0].ToString(), (double)row.ToArray()[1], (double)row.ToArray()[2]));
                }
            }

            foreach (Language l in languageList)
                Console.WriteLine(l.ToString());
        }
    }
}
