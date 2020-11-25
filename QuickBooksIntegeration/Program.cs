namespace QuickBooksIntegeration
{
    using MySql.Data.MySqlClient;
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Data;
    using System.Diagnostics;
    using System.IO;

    /// <summary>
    /// Defines the <see cref="Program" />.
    /// </summary>
    internal class Program
    {
        /// <summary>
        /// Defines the connectionString.
        /// </summary>
        public static string connectionString;

        /// <summary>
        /// Defines the dbConnection.
        /// </summary>
        public static MySqlConnection dbConnection;

        /// <summary>
        /// The Main.
        /// </summary>
        /// <param name="args">The args<see cref="string[]"/>.</param>
        internal static void Main(string[] args)
        {
            Console.WriteLine("\nSilver Solve's QuickBooks Data Fetching Agent At Service: \n");
            Console.WriteLine("*********************************************************************** \n");
            while (true)
            {
                // CREATE LOGS FOLDER IF NOT EXISTS:
                System.IO.Directory.CreateDirectory(ConfigurationManager.AppSettings["Logs_Folder"]);


                using (StreamWriter logsFile = new StreamWriter(ConfigurationManager.AppSettings["Logs_Folder"] + "/QB_Agent_" + DateTime.Now.ToString("dd_MM_yyyy") + ".txt", true))
                {

                    try
                    {
                        // DATABASE CONNECTION:
                        MakeDatabaseConnection(logsFile);

                        // FETCH ALL COMPANIES TO SYNC:
                        List<dynamic> allCompanies = FetcPendingCompanies(logsFile);
                        var filePath = "";
                        var companyName = "";
                        var companyFolderName = "";
                        var companyFile = "";
                        var companyId = "";
                        var sync_start_time = "";
                        var sync_end_time = "";
                        var current_date_time = DateTime.Now.ToString("HH:mm:ss");
                        var allocatedSyncStartTime = "18:00:00";
                        var allocatedSyncEndTime = "18:30:00";
                        Stopwatch stopWatch = new Stopwatch();


                        // HOUR IN WHICH QB AGENT WILL FETCHED DATA:
                        if (allCompanies.Count > 0)
                        {
                            allocatedSyncStartTime = allCompanies[0][4];
                            allocatedSyncEndTime = allCompanies[0][5];
                        }



                        //CloseQuickBooks(logsFile);
                        //.AddMinutes(300)
                        if (DateTime.Parse(current_date_time) >= DateTime.Parse(allocatedSyncStartTime) & DateTime.Parse(current_date_time) <= DateTime.Parse(allocatedSyncEndTime))
                        {
                            CloseQuickBooks(logsFile);
                            //Console.WriteLine("TEST\t" + allocatedSyncStartTime);

                        }

                        // ITERATING OVER ALL COMPANIES AND FETCHING AND INSERTING QB DATA:
                        for (int i = 0; i < allCompanies.Count; i++)
                        {
                            
                            Console.WriteLine("\n");
                            logsFile.WriteLine("\t\t\t************************************Silver Solve's QuickBooks Data Fetching Agent Log Service***********************************************");


                            //logsFile.WriteLine("CONNECTION STATUS  AT START......\n" + SessionManager.Instance.IsConnectionOpen);
                            //logsFile.WriteLine("SessionManager STATUS AT START......\n" + SessionManager.Instance.IsSessionBegun);


                            Console.WriteLine("Total Number Of Companies To Sync:\t" + allCompanies.Count);
                            logsFile.WriteLine();
                            logsFile.WriteLine(DateTime.Now + "\tINFO\tTotal Number Of Companies To Sync   :" + allCompanies.Count);
                            logsFile.WriteLine();
                            Console.WriteLine("\n");

                            try
                            {
                                CloseQuickBooks(logsFile);
                                companyId = allCompanies[i][0];
                                companyName = allCompanies[i][1];
                                companyFolderName = allCompanies[i][2];
                                sync_start_time = DateTime.Now.ToString("HH:mm:ss");
                                filePath = ConfigurationManager.AppSettings["Path"] + companyFolderName;
                                string[] companyFileDir = Directory.GetFiles(filePath, "*.QBW");
                                if (companyFileDir.Length > 0)
                                {
                                    // GET COMPANY FILE:
                                    companyFile = companyFileDir[0];
                                    Console.WriteLine("Fetching GL Accounts, Customer & Vendor Invoices Data of " + companyName + " & Inserting It In To Database Please Wait...........\n");
                                    logsFile.WriteLine(DateTime.Now + "\tINFO\tFetching GL Accounts, Customer & Vendor Invoices Data of " + companyName + " & Inserting It In To Database Please Wait...........");
                                    logsFile.WriteLine();
                                    Console.WriteLine("Synchronization Process Started At:\t" + sync_start_time + "\n");
                                    logsFile.WriteLine(DateTime.Now + "\tINFO\tSynchronization Process Started At:\t" + sync_start_time + "");
                                    logsFile.WriteLine();
                                    Console.WriteLine("Execution In Progres......\n");
                                    logsFile.WriteLine(DateTime.Now + "\tINFO\tExecution In Progres......");
                                    logsFile.WriteLine();
                                    var updateCompanySynInfo = "UPDATE companies_data SET sync_process_started='" + sync_start_time + "',sync_status='In Progress' WHERE id=" + companyId + "";
                                    connectionString = ConfigurationManager.ConnectionStrings["qbAgentMySqlDbConnection"].ConnectionString;
                                    var Command = new MySqlCommand(updateCompanySynInfo, dbConnection);
                                    Command.ExecuteNonQuery();
                                    SessionManager.Instance.IsSessionBegun = false;
                                    SessionManager.Instance.IsConnectionOpen = false;
                                    Query query = new Query();
                                    logsFile.WriteLine("CONNECTION STATUS......\n" + SessionManager.Instance.IsConnectionOpen);
                                    logsFile.WriteLine("SessionManager STATUS......\n" + SessionManager.Instance.IsSessionBegun);

                                    //*******************************************************

                                    SessionManager.Instance.OpenConnection(companyName, logsFile);

                                    if (SessionManager.Instance.IsConnectionOpen == true)
                                    {
                                        SessionManager.Instance.BeginSession(companyFile, logsFile);

                                        if (SessionManager.Instance.IsSessionBegun == true)
                                        {
                                            //CUSTOMER DATA:

                                            Console.WriteLine("Fetching Customers Data of " + companyName + " Company\n");
                                            logsFile.WriteLine(DateTime.Now + "\tINFO\tFetching Customers Data of " + companyName + " Company");
                                            logsFile.WriteLine();
                                            stopWatch.Start();
                                            List<dynamic> customersData = query.GetCustomersData(logsFile);
                                            for (int j = 0; j < customersData.Count; j++)
                                            {
                                                var insertCustomersDataQuery = "INSERT INTO customer_query(company_id,customer_name,account_number,total_balance,date_time) VALUES('" + companyId + "','" + MySql.Data.MySqlClient.MySqlHelper.EscapeString(customersData[j][0]) + "','" + MySql.Data.MySqlClient.MySqlHelper.EscapeString(customersData[j][1]) + "','" + MySql.Data.MySqlClient.MySqlHelper.EscapeString(customersData[j][2]) + "',CURDATE() - INTERVAL 1 DAY)";
                                                Command = new MySqlCommand(insertCustomersDataQuery, dbConnection);
                                                Command.ExecuteNonQuery();
                                            }

                                            //*******************************************************

                                            //VENDOR DATA:

                                            Console.WriteLine("Fetching Vendors Data of " + companyName + " Company\n");
                                            logsFile.WriteLine(DateTime.Now + "\tINFO\tFetching Vendors Data of " + companyName + " Company");
                                            logsFile.WriteLine();
                                            List<dynamic> vendorsData = query.GetVendorsData(logsFile);
                                            for (int j = 0; j < vendorsData.Count; j++)
                                            {
                                                var insertVendorsDataQuery = "INSERT INTO vendor_query(company_id,vendor_name,account_number,vendor_balance,date_time) VALUES('" + companyId + "','" + MySql.Data.MySqlClient.MySqlHelper.EscapeString(vendorsData[j][0]) + "','" + MySql.Data.MySqlClient.MySqlHelper.EscapeString(vendorsData[j][1]) + "','" + MySql.Data.MySqlClient.MySqlHelper.EscapeString(vendorsData[j][2]) + "',CURDATE() - INTERVAL 1 DAY)";
                                                Command = new MySqlCommand(insertVendorsDataQuery, dbConnection);
                                                Command.ExecuteNonQuery();
                                            }

                                            //*******************************************************

                                            //ACCOUNT DATA:

                                            Console.WriteLine("Fetching Accounts Data of " + companyName + " Company\n");
                                            logsFile.WriteLine(DateTime.Now + "\tINFO\tFetching Accounts Data of " + companyName + " Company");
                                            logsFile.WriteLine();
                                            List<dynamic> accountsData = query.GetQBAccounts(logsFile);
                                            for (int j = 0; j < accountsData.Count; j++)
                                            {
                                                var insertAccountsDataQuery = "INSERT INTO account_details(company_id,account_name,account_number,account_type,total_balance,date_time) VALUES('" + companyId + "','" + MySql.Data.MySqlClient.MySqlHelper.EscapeString(accountsData[j][0]) + "','" + MySql.Data.MySqlClient.MySqlHelper.EscapeString(accountsData[j][1]) + "','" + MySql.Data.MySqlClient.MySqlHelper.EscapeString(accountsData[j][2]) + "','" + MySql.Data.MySqlClient.MySqlHelper.EscapeString(accountsData[j][3]) + "',CURDATE() - INTERVAL 1 DAY)";
                                                Command = new MySqlCommand(insertAccountsDataQuery, dbConnection);
                                                Command.ExecuteNonQuery();
                                            }

                                            sync_end_time = DateTime.Now.ToString("HH:mm:ss");
                                            Console.WriteLine("Accounts, Customers & Vendors Data of " + companyName + " Company Fetched Successfully\n");
                                            logsFile.WriteLine(DateTime.Now + "\tINFO\tAccounts, Customers & Vendors Data of " + companyName + " Company Fetched Successfully");
                                            logsFile.WriteLine();
                                            Console.WriteLine("Operation Successfull...........\n");
                                            logsFile.WriteLine(DateTime.Now + "\tINFO\tOperation Successfull...........");
                                            logsFile.WriteLine();
                                            Console.WriteLine("Synchronization Process Ended At:\t" + sync_end_time);
                                            logsFile.WriteLine(DateTime.Now + "\tINFO\tSynchronization Process Ended At:" + sync_end_time);
                                            logsFile.WriteLine();
                                            Console.WriteLine("\n");
                                            stopWatch.Stop();
                                            TimeSpan total_time = stopWatch.Elapsed;
                                            Console.WriteLine("Total Time Synchronization Process Took:\t" + total_time);
                                            logsFile.WriteLine(DateTime.Now + "\tINFO\tTotal Time Synchronization Process Took:" + total_time);
                                            logsFile.WriteLine();
                                            Console.WriteLine("\n");
                                            SessionManager.Instance.EndSession(logsFile);
                                            SessionManager.Instance.CloseConnection(logsFile);
                                            updateCompanySynInfo = "UPDATE companies_data SET sync_start_date=DATE_ADD(sysdate(), INTERVAL 1 DAY) ,sync_process_ended='" + sync_end_time + "', sync_total_time='" + total_time + "',sync_status='Done' WHERE id=" + companyId + "";
                                            Command = new MySqlCommand(updateCompanySynInfo, dbConnection);
                                            Command.ExecuteNonQuery();
                                            stopWatch.Reset();

                                        }
                                        else
                                        {
                                            logsFile.WriteLine(DateTime.Now + "\tWARNING\tQB Error Establishing Session\n");
                                            Console.WriteLine("WARNING:  Error Establishing Session\n");
                                            SessionManager.Instance.EndSession(logsFile);
                                            SessionManager.Instance.CloseConnection(logsFile);
                                        }
                                    }
                                    else
                                    {
                                        logsFile.WriteLine(DateTime.Now + "\tWARNING\tQB Error Establishing Connection\n");
                                        Console.WriteLine("WARNING:  Error Establishing Connection\n");
                                        SessionManager.Instance.CloseConnection(logsFile);
                                    }
                                }
                                else
                                {
                                    logsFile.WriteLine(DateTime.Now + "\tWARNING\tQB COMPANY FILE DOES NOT EXISTS\n");
                                    Console.WriteLine("WARNING:  QB COMPANY FILE DOES NOT EXISTS\n");
                                }

                            }
                            catch (Exception ex)
                            {
                                var updateSyncStatus = "UPDATE companies_data SET sync_status ='Pending' WHERE id=" + companyId + "";
                                var Command = new MySqlCommand(updateSyncStatus, dbConnection);
                                Command.ExecuteNonQuery();
                                Console.WriteLine("Warning Error Occured During Execution, Please Refer to Logs...\n");
                                logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\t" + ex.Message);
                                logsFile.WriteLine();
                                logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\tStack Trace:\t" + ex.StackTrace);
                                logsFile.WriteLine();
                                SessionManager.Instance.EndSession(logsFile);
                                SessionManager.Instance.CloseConnection(logsFile);
                                CloseQuickBooks(logsFile);

                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Warning Error Occured During Execution, Please Refer to Logs...\n");
                        logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\t" + ex.Message);
                        logsFile.WriteLine();
                        logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\tStack Trace:\t" + ex.StackTrace);
                        logsFile.WriteLine();
                    }
                }


            }
        }

        /// <summary>
        /// The MakeDatabaseConnection.
        /// </summary>
        /// <param name="logsFile">The logsFile<see cref="StreamWriter"/>.</param>
        internal static void MakeDatabaseConnection(StreamWriter logsFile)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["qbAgentMySqlDbConnection"].ConnectionString;
                dbConnection = new MySqlConnection(connectionString);
                //Console.WriteLine("Database Connection In Progress...\n");
                if (dbConnection != null && dbConnection.State == ConnectionState.Closed)
                {
                    MySqlConnection.ClearPool(dbConnection);
                    dbConnection.Open();
                    //Console.WriteLine("Database Connected Successfully...\n");
                }
                else
                {
                    Console.WriteLine("Database Connection Failed, Please Contact Silver Solve...\n");
                }

            }
            catch (Exception ex)
            {
                //StreamWriter logsFile = new StreamWriter("Logs/QB_Agent_" + DateTime.Now.ToString("dd_MM_yyyy") + ".txt", true);


                Console.WriteLine("Warning Error Occured During Execution, Please Refer to Logs...\n");
                logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\t" + ex.Message);
                logsFile.WriteLine();
                logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\tStack Trace:\t" + ex.StackTrace);
                logsFile.WriteLine();


            }
        }

        /// <summary>
        /// The FetcPendingCompanies.
        /// </summary>
        /// <param name="logsFile">The logsFile<see cref="StreamWriter"/>.</param>
        /// <returns>The <see cref="List{dynamic}"/>.</returns>
        internal static List<dynamic> FetcPendingCompanies(StreamWriter logsFile)
        {
            List<dynamic> allCompanies = new List<dynamic>();
            try
            {

                var getCompanyQuery = "SELECT * FROM companies_data  WHERE Date(sync_start_date) <= Date(sysdate()) AND sync_start_time <= Time(sysdate())";
                var Command = new MySqlCommand(getCompanyQuery, dbConnection);
                MySqlDataReader DataReader = Command.ExecuteReader();
                while (DataReader.Read())
                {
                    allCompanies.Add(new string[] { Convert.ToString(DataReader["id"]), Convert.ToString(DataReader["company_name"]), Convert.ToString(DataReader["company_folder_name"]), Convert.ToString(DataReader["sync_start_date"]), Convert.ToString(DataReader["sync_start_time"]), Convert.ToString(DataReader["sync_endt_time"]) });
                }
                DataReader.Close();
                if (allCompanies.Count == 0)
                {
                    var updateSyncStatus = "UPDATE companies_data SET sync_status ='Pending'";
                    var syncStatus = new MySqlCommand(updateSyncStatus, dbConnection);
                    //Console.WriteLine("All Quickbooks Companies Are Syncronized, Therefore Aborting Synchronization For Today\n...");
                    syncStatus.ExecuteNonQuery();
                    //logsFile.Close();
                }
                //else
                //{
                //    Console.WriteLine("Fetching Quickbooks Comapanies For Data Synchronization...\n");
                //    Console.WriteLine("Companies Fetched Successfully...");
                //}
                return allCompanies;

            }
            catch (Exception ex)
            {
                Console.WriteLine("Warning Error Occured During Execution, Please Refer to Logs...\n" + ex.Message);
                logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\t" + ex.Message);
                logsFile.WriteLine();
                logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\tStack Trace:\t" + ex.StackTrace);
                logsFile.WriteLine();
                return allCompanies;
            }
        }

        /// <summary>
        /// The CloseQuickBooks.
        /// </summary>
        /// <param name="logsFile">The logsFile<see cref="StreamWriter"/>.</param>
        public static void CloseQuickBooks(StreamWriter logsFile)
        {
            try
            {
                foreach (Process Proc in Process.GetProcessesByName("QBW32"))
                    Proc.Kill();
                
            }
            catch (Exception ex)
            {
                Console.WriteLine("Warning Error Occured During Execution, Please Refer to Logs...\n" + ex.Message);
                logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\t" + ex.Message);
                logsFile.WriteLine();
                logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\tStack Trace:\t" + ex.StackTrace);
                logsFile.WriteLine();

            }
        }
    }
}
