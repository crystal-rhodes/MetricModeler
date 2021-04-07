using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace MetricModeler
{
    public partial class Form1 : Form
    {
        List<ProjectHistory> projectHistoryList;
        List<Language> languageList;
        const double EAF = 1; // effort adjustment factor
        const double T = 0.35; // sloc-dependent coefficient
        double P = 1.14; // project complexity
        const double pricingPerHour = 50.00;

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

        private void Form1_Load(object sender, EventArgs e)
        {
            languageComboBox.DataSource = languageList.Select(l => l.LanguageName).ToList();

            List<string> projectTypes = new List<string>();
            projectTypes.Add("1 - Linear");
            projectTypes.Add("2 - OO");
            projectTypes.Add("3 - Other");
            projectTypeComboBox.DataSource = projectTypes;

            softwareDevelopmentCapabilityTextBox.Text = "3";

            List<string> functionExpectationList = new List<string>();
            functionExpectationList.Add("1 - Very low");
            functionExpectationList.Add("2 - Low");
            functionExpectationList.Add("3 - Medium");
            functionExpectationList.Add("4 - High");
            functionExpectationList.Add("5 - Very high");

            functionExpectationComboBox.DataSource = functionExpectationList;
            functionExpectationComboBox.SelectedIndex = 2;

            designReviewHoursTextBox.Text = "8";
            averageCostPerHourTextBox.Text = "40";
            numTables.Text = "10";

            noInputWeightingFactor.Text = "4";
            noOutputWeightingFactor.Text = "5";
            noInquiriesWeightingFactor.Text = "4";
            noLogicalFilesWeightingFactor.Text = "10";
            noExternalInterfacesWeightingFactor.Text = "7";
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
                                (int)reader["LOC"],
                                (int)reader["Estimated FP"],
                                (int)reader["Actual FP"],
                                (int)reader["Ave Cost per Person Hour"],
                                (int)reader["Software Development Capability"],
                                (int)reader["Design Review Hours"],
                                reader["Development Language"].ToString(),
                                (int)reader["Language Productivity Factor"],
                                (int)reader["Required Functionalities Expectation"],
                                (int)reader["Number of Tables"]

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

        public static readonly int scale = 3;

        private void calculateButton_Click(object sender, EventArgs e)
        {
            try
            {
                int factor = 14 * scale;
                double complexitiyAdjustmentFactor = 0.65 + (0.01 * factor);

                int softwareDevelopmentCapability = int.Parse(softwareDevelopmentCapabilityTextBox.Text);
                int averageCostPerHour = int.Parse(averageCostPerHourTextBox.Text);
                int designReviewHours = int.Parse(this.designReviewHoursTextBox.Text);
                int NumTables = int.Parse(this.numTables.Text);

                int inputsValue = int.Parse(noInputValue.Text);
                int outputsValue = int.Parse(noOutputValue.Text);
                int inquiriesValue = int.Parse(noInquiriesValue.Text);
                int logicalFilesValue = int.Parse(noLogicalFilesValue.Text);
                int externalInterfacesValue = int.Parse(noExternalInterfacesValue.Text);

                int inputsWeightingFactor = int.Parse(noInputWeightingFactor.Text);
                int outputsWeightingFactor = int.Parse(noOutputWeightingFactor.Text);
                int inquiriesWeightingFactor = int.Parse(noInquiriesWeightingFactor.Text);
                int logicalFilesWeightingFactor = int.Parse(noLogicalFilesWeightingFactor.Text);
                int externalInterfacesWeightingFactor = int.Parse(noExternalInterfacesWeightingFactor.Text);

                if (softwareDevelopmentCapability >= 10)
                {
                    statusLabel.Text = "Software Development Capability must be < 10";
                    return;
                }

                // Calculate unadjusted function points
                int unadjustedFunctionPoints =
                            (inputsValue * inputsWeightingFactor) +
                            (outputsValue * outputsWeightingFactor) +
                            (inquiriesValue * inquiriesWeightingFactor) +
                            (logicalFilesValue * logicalFilesWeightingFactor) +
                            (externalInterfacesValue * externalInterfacesWeightingFactor);

                // Calculate function points 
                double functionPoints = complexitiyAdjustmentFactor * unadjustedFunctionPoints;

                // Get Language Productivity Factor
                double languageProductivityFactor = languageList.Find(l => l.LanguageName == languageComboBox.SelectedItem.ToString()).Level;

                // Calculate lines of code
                double LOC = functionPoints * languageProductivityFactor;

                // Calculate thousands of code
                double KLOC = functionPoints * languageProductivityFactor / 1000;

                // Calculate required functionalities Expectation
                int functionLevel = int.Parse(functionExpectationComboBox.SelectedItem.ToString()[0].ToString());
                Console.WriteLine(functionLevel);
                switch (functionLevel)
                {
                    case 1:
                        P = 1.04;
                        break;
                    case 2:
                        P = 1.09;
                        break;
                    case 3:
                        P = 1.14;
                        break;
                    case 4:
                        P = 1.19;
                        break;
                    case 5:
                        P = 1.24;
                        break;
                    default:
                        P = 1.14;
                        break;
                }

                // PM = 2.45*EAF*(SLOC/1000)^P
                double personMonth = 2.45 * EAF * Math.Pow(LOC / 100, P) * (NumTables * 0.01);

                personMonth = personMonth * (100 - (1.0 * softwareDevelopmentCapability / 5 * 10)) / 100;

                // DM = 2.50*(PM)^T
                double durationMonths = 2.50 * Math.Pow(personMonth, T);

                // Assuming that 7 working hours per day, 20 days per month, and 12 working months per year
                double durationDays = durationMonths * 20 * 7;

                double designReviewCost = designReviewHours * pricingPerHour;

                double cost = designReviewCost + (durationDays * averageCostPerHour);

                timeLabel.Text = Math.Round(durationMonths, 2) + " Months"; // time
                scopeLabel.Text = Math.Round(personMonth, 2) + " Person-months"; // scope
                costLabel.Text = Math.Round(cost, 2) + "$";
                functionPointsLabel.Text = Math.Round(functionPoints, 2).ToString();
                klocLabel.Text = Math.Round(KLOC, 2).ToString();
                languageProductivityLabel.Text = languageProductivityFactor.ToString();

                statusLabel.Text = "Size metrics of the project have been calculated. Waiting for new input";



                // BELOW THIS IS THE STUFF I ADDED TO SORT THE TABLES

                Console.WriteLine("Part 4 Compare projects based on the cost, scope and  time");
                Console.WriteLine("");

                Console.WriteLine("Sorted By Cost");

                List<ProjectHistory> copyList = projectHistoryList.OrderBy(o => o.EstProjectCost).ToList();


                foreach (ProjectHistory p in copyList)
                {

                    Console.WriteLine("Name: " + p.ProjectName + " The type of project: " + p.ProjectType + "Estimated Cost" + p.EstProjectCost + "Actual Cost" + p.ActualProjectCost);

                }

                Console.WriteLine("Sorted By Scope");

                copyList = projectHistoryList.OrderBy(o => o.EstimatedFP).ToList();


                foreach (ProjectHistory p in copyList)
                {

                    Console.WriteLine("Name: " + p.ProjectName + " The type of project: " + p.ProjectType + "Estimated Scope" + p.EstimatedFP + "Actual Scope" + p.ActualFP);

                }

                Console.WriteLine("Sorted By Time");

                copyList = projectHistoryList.OrderBy(o => o.EstDuration).ToList();

                TimeSpan ts;
                int diffDays;

                foreach (ProjectHistory p in copyList)
                {
                    ts = p.EndDate - p.StartDate;
                    diffDays = ts.Days;
                    Console.WriteLine("Name: " + p.ProjectName + " The type of project: " + p.ProjectType + "Estimated Duration" + p.EstDuration + "Actual Time" + diffDays);

                }


            }
            catch (Exception ex)
            {
                statusLabel.Text = "Some fields are missing or are not entered in correct format.";
            }
        }

        private void label23_Click(object sender, EventArgs e)
        {

        }
    }
}
