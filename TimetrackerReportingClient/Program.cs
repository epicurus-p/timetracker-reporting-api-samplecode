using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using CommandLine;
using ExtendedXmlSerializer.Configuration;
using Newtonsoft.Json;
using RestSharp.Serializers;
using TimetrackerOnline.Reporting.Models;
using Formatting = Newtonsoft.Json.Formatting;
using XmlSerializer = System.Xml.Serialization.XmlSerializer;

namespace TimetrackerReportingClient
{
    internal class Program
    {
        const String employeePath = @".\Data\employee.xml";
        const String disciplinePath = @".\Data\WBSDiscipline.xml";
        const String rndPath = @".\Data\WBSRandD.xml";

        static XmlDocument employee = new XmlDocument();
        static XmlDocument wbsDiscipline = new XmlDocument();
        static XmlDocument wbsRandD = new XmlDocument();

        private static void Main(string[] args)
        {
            bool parsed = false;
            CommandLineOptions cmd = null;
            // Get parameters
            CommandLine.Parser.Default.ParseArguments<CommandLineOptions>(args).WithParsed(x =>
            {
                parsed = true;
                cmd = x;
            })
                       .WithNotParsed(x => { Console.WriteLine("Check https://github.com/7pace/timetracker-reporting-api-samplecode to get samples of usage"); });

            if (!parsed)
            {
                Console.ReadLine();
                return;
            }
            DateTime endstdt;

            Console.WriteLine("Enter Sunday Week Ending Date (YYYY/MM/DD):");
            string endDate = Console.ReadLine(); 

            if (DateTime.TryParse(endDate, out endstdt))
            {
                if (endstdt.DayOfWeek != DayOfWeek.Sunday)
                {
                    Console.WriteLine("Week Ending Date must be a Sunday");
                    System.Environment.Exit(0);
                }                        
            }    
            else
            {
                Console.WriteLine("Could not parse Week Ending Date");
                System.Environment.Exit(0);
            }

            // Create OData service context
            var context = cmd.IsWindowsAuth
                ? new TimetrackerOdataContext(cmd.ServiceUri)
                : new TimetrackerOdataContext(cmd.ServiceUri, cmd.Token);

            Console.WriteLine("Querying Timesheets...");
            // request for work items with worklogs
            var workLogsWorkItemsExport = context.Container.workLogsWorkItems;
            //fills custom fields values if provided. Check https://support.7pace.com/hc/en-us/articles/360035502332-Reporting-API-Overview#user-content-customfields to get more information
            if (cmd.CustomFields != null && cmd.CustomFields.Any())
            {
                workLogsWorkItemsExport = workLogsWorkItemsExport.AddQueryOption("customFields", string.Join(",", cmd.CustomFields));
            }
            var workLogsWorkItemsExportResult = workLogsWorkItemsExport
                // Perform query for 3 last months
                .Where(s => s.Timestamp > endstdt.AddDays(-7) && s.Timestamp < endstdt)
                // orfer items by worklog date
                .OrderByDescending(g => g.WorklogDate.ShortDate).ToArray();

            DataTable timesheetData = TimesheetData.CreateDataTable();

            // Print out the result
            foreach (var row in workLogsWorkItemsExportResult)
            {
                if (row.WorkItem != null)
                {
                    double periodLength = row.PeriodLength;
                    double hours = periodLength / 3600;

                    Console.WriteLine("{0:g} {1} {2} {3} {4} {5} {6}", row.WorklogDate.ShortDate, row.User.Name, hours, row.WorkItem.Microsoft_VSTS_Common_Activity, row.WorkItem.System_Id, row.WorkItem.System_WorkItemType, row.WorkItem.CustomStringField1);
                                   
                    if ((row.WorkItem.System_WorkItemType == "Bug") || (row.WorkItem.System_WorkItemType == "Task") || (row.WorkItem.System_WorkItemType == "Bug Backfix"))
                    {
                        DataRow timesheetRow = timesheetData.NewRow();

                        timesheetRow["User"] = row.User.Email.Split('@')[0];
                        timesheetRow["Date"] = row.WorklogDate.ShortDate;

                        timesheetRow["Hours"] = String.Format("{0:0.##}", hours);
                        timesheetRow["Discipline"] = row.WorkItem.Microsoft_VSTS_Common_Activity;
                        timesheetRow["RnD Type"] = "0";
                        timesheetRow["Job Code"] = row.WorkItem.CustomStringField1;
                        timesheetRow["Work Item Id"] = row.WorkItem.System_Id.ToString();
                        timesheetRow["Project"] = row.WorkItem.System_TeamProject;

                        timesheetData.Rows.Add(timesheetRow);
                    }
                    else
                    {
                        string strLogText = String.Format("Invalid Work Item Type'{0}' logged by user {1}\n", row.WorkItem.System_WorkItemType, row.User.Name);
                        Console.WriteLine(strLogText);
                    }
                }
                else
                {
                    string strLogText = String.Format("No Work Item specified by user {0}\n", row.User.Name);
                    Console.WriteLine(strLogText);
                }
            }

            //Export(cmd.Format, workLogsWorkItemsExportResult, "workLogsWorkItemsExport");

            Console.WriteLine("\r\nTimesheets Processed.");

            ExportToSAP(timesheetData);

            Console.WriteLine("\r\nTimesheets Exported.");
        }

        public static void Export(string format, object extendedData, string fileName)
        {
            if (string.IsNullOrEmpty(format))
            {
                return;
            }

            //save here
            string location = System.Reflection.Assembly.GetExecutingAssembly().Location;

            //once you have the path you get the directory with:
            var directory = System.IO.Path.GetDirectoryName(location);

            if (format == "xml")
            {
                var serializer = new ConfigurationContainer()

                    // Configure...
                    .Create();

                var exportPath = directory + $"/{fileName}.xml";

                var file = File.OpenWrite(exportPath);
                var settings = new XmlWriterSettings { Indent = true };

                var xmlTextWriter = new XmlTextWriter(file, Encoding.UTF8);
                xmlTextWriter.Formatting = System.Xml.Formatting.Indented;

                xmlTextWriter.Indentation = 4;

                serializer.Serialize(xmlTextWriter, extendedData);
                xmlTextWriter.Close();
                xmlTextWriter.Dispose();
                file.Close();
                file.Dispose();
            }
            else if (format == "json")
            {
                var json = JsonConvert.SerializeObject(extendedData, Formatting.Indented);
                var exportPath = directory + $"/{fileName}.json";
                File.WriteAllText(exportPath, json);
            }
            else
            {
                throw new NotSupportedException("Provided format is not supported: " + format);
            }

            Console.WriteLine($"\r\nExport to file {fileName}.{format} completed");
        }

        public static void ExportToSAP(DataTable timesheetData)
        {
            DateTime today = DateTime.Today;
            string csvpath = string.Empty;

            //Load XML file for mapping process
            employee.Load(employeePath);
            wbsDiscipline.Load(disciplinePath);
            wbsRandD.Load(rndPath);

            List<DataTable> finalConvertedData = new List<DataTable>();

            CsvParser.CreateDataTable create = new CsvParser.CreateDataTable();
            DataTable dataTable = create.CreatingDataTable();

            int count = 0;
            int countLimit = 200;

            foreach (DataRow dataRow in timesheetData.Rows)
            {
                DataRow row = dataTable.NewRow();

                string userName = dataRow["User"].ToString();
                string date = dataRow["Date"].ToString();
                string hours = dataRow["Hours"].ToString();
                string discipline = dataRow["Discipline"].ToString();
                string rnd = dataRow["RnD Type"].ToString();
                string jobCode = dataRow["Job Code"].ToString();
                string workID = dataRow["Work Item Id"].ToString();
                string project = dataRow["Project"].ToString();

                if (string.IsNullOrEmpty(userName))
                {
                    throw new ArgumentNullException();
                }
                else
                {
                    //get employee ID
                    string empID = GetEmployeeCode(userName, employee);
                    if (empID != null)
                    {
                        row["Employee Number"] = empID;
                    }
                    else if (empID == null)
                    {
                      //  string strLogText = String.Format("'{0}' is not in the employeecode.xml\n", userName);
                      //  Console.WriteLine(strLogText);
                    }

                    row["Work Date"] = DateTime.Parse(date).ToString("dd-MMM-yy");

                    DateTime weekEnding = DateTime.Parse(GetWeekEnding(date));
                    today = weekEnding;

                    row["Week Ending"] = weekEnding.ToString("dd-MMM-yy");
                    row["Hours"] = hours;

                    //start creating/sorting SAP wbs level
                    string sapWBSNo = GetWBSLevel(jobCode, discipline, rnd, userName, workID, project);
                    row["SAP WBS No."] = sapWBSNo;

                    if (sapWBSNo != "" && sapWBSNo != null && empID != null)
                    {
                        dataTable.Rows.Add(row);
                    }
                }

                if (count++ == countLimit)
                {
                    finalConvertedData.Add(dataTable);
                    dataTable = create.CreatingDataTable();
                    count = 0;
                }
           
            }

            finalConvertedData.Add(dataTable);

            //Parse the data to csv file - SAP format
            CsvParser.CsvWriter csvWriter = new CsvParser.CsvWriter();
            int add = 1;

            DateTime weekEnding2 = today;
            string endDate_FileNameNew = (string.Format("{0:yyyyMMdd}", weekEnding2));

            DateTime addDays = weekEnding2.AddDays(-6);
            string startDate_FileNameNew = (string.Format("{0:yyyyMMdd}", addDays));

            //Joining all the codes - SAP WBS No.
            string[] join = { startDate_FileNameNew, "_", endDate_FileNameNew, "_", "(", add.ToString(), ")", ".", "csv" };
            string fullFileName = String.Join("", join);

            //Write the basedata out as well
            StreamWriter writer2 = new StreamWriter("base_timesheetData" + fullFileName);
            csvWriter.WriteToStream(writer2, timesheetData, true, false);

            foreach (DataTable tb in finalConvertedData)
            {
                //Joining all the codes - SAP WBS No.
                 string[] tjoin = { startDate_FileNameNew, "_", endDate_FileNameNew, "_", "(", add.ToString(), ")", ".", "csv" };
                 fullFileName = String.Join("", tjoin);

                //Write the SAP file to specific file and location
                StreamWriter writer = new StreamWriter(fullFileName);
                csvWriter.WriteToStream(writer, tb, true, false);

                writer.Close();
                writer.Dispose();

                add++;
            }

            writer2.Close();
            writer2.Dispose();

        }
        private static string GetEmployeeCode(string name, XmlDocument employee)
        {
            string Id = null;

            XmlNode node = employee.SelectSingleNode(@"root/employee[translate(name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz') = '" + name.ToLower() + "']/id");
            if (node != null)
            {
                Id = node.InnerText;
            }
            else if (node == null)
            {
                Id = null;
            }

            return Id;
        }

        /// <summary>
        /// Get week ending for working date
        /// </summary>
        /// <returns>Week Ending</returns>
        private static string GetWeekEnding(string date)
        {
            string endValue = null;
            IFormatProvider format = new CultureInfo("en-AU");
            DateTime dateTime = DateTime.Parse(date, format);
            DayOfWeek day = (DateTime.Parse(dateTime.ToString()).DayOfWeek);
            if (day.ToString() == "Sunday")
            {
                int days = day - DayOfWeek.Monday;
                DateTime start = DateTime.Parse(dateTime.ToString()).AddDays(-days);
                DateTime end = start.AddDays(-1);
                endValue = end.ToString();
            }
            else
            {
                int days = day - DayOfWeek.Monday;
                DateTime start = DateTime.Parse(dateTime.ToString()).AddDays(-days);
                DateTime end = start.AddDays(6);
                endValue = end.ToString();
            }

            return endValue;
        }

        /// <summary>
        /// Search for Job Code, discipline, rnd or admin level and returns SAP code.
        /// </summary>
        /// <returns>SAP code</returns>
        private static string GetWBSLevel(string jobCode, string discipline, string rnd, string name, string workid, string project)
        {
            string node = null;

            //Get employee location from EmployeeCodes.xml
            string employeelocation = GetEmployeeLocation(name);

            //Get SapCode
            if (project == "Time Tracking")
            {
                if (jobCode.StartsWith("RLV") || jobCode.StartsWith("RAD"))
                {
                    node = jobCode;

                    node = node.Replace("XX", employeelocation);
                }
                else
                {
                    string strLogText = String.Format("Job Code for work item '{0}' is invalid\n", workid);
                    Console.WriteLine(strLogText);
                }
            }
            else
            {
                int index = jobCode.IndexOf('(');
                if (index > 0)
                {
                    node = jobCode.Substring(index+1, jobCode.Length - index - 2);
                    node = node + "-01-01-" + GetDisciplineCode(discipline, workid);
                }
                else
                {
                    node = String.Empty;
                    string strLogText = String.Format("Job Code for work item '{0}' is invalid\n", workid);
                    Console.WriteLine(strLogText);
                }               
            }

            return node;
        }


        /// <summary>
        /// Get employee location from EmployeeCodes.xml
        /// </summary>
        /// <returns>Employee Location</returns>
        private static string GetEmployeeLocation(string name)
        {
            string EmpLoc = null;
            XmlNode node = employee.SelectSingleNode(@"root/employee[translate(name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz') = '" + name.ToLower() + "']/location");

            if (node != null)
            {
                EmpLoc = node.InnerText;
            }
            else if (node == null)
            {
                EmpLoc = null;
                string strLogText = String.Format("Could not find employee '{0}' in the employee list.\n", name);
                Console.WriteLine(strLogText);
            }

            return EmpLoc;
        }

        /// <summary>
        /// Get employee location from EmployeeCodes.xml
        /// </summary>
        /// <returns>Employee Location</returns>
        private static string GetDisciplineCode(string discipline, string workId)
        {
            string disciplineCode = null;
            XmlNode node = wbsDiscipline.SelectSingleNode(@"root/discipline[translate(name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz') = '" + discipline.ToLower() + "']/WBS");

            if (node != null)
            {
                disciplineCode = node.InnerText;
            }
            else if (node == null)
            {
                disciplineCode = null;
                string strLogText = String.Format("Activity '{0}' for work item '{1}' is not in the Activity List file.\n", discipline, workId);
                Console.WriteLine(strLogText);
            }

            return disciplineCode;
        }

    }
}