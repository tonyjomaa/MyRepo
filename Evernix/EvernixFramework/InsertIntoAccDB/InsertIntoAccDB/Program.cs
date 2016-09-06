// ********************************************************************************
//
// Copyright (c) 2013, Algorithmic Implementations, Inc. (dba Ai Squared). All Rights Reserved.
//
// Author: Tony Jomaa
//
// Description: Insert a new row into Access DB
//
// ********************************************************************************
using System;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Text;
//using System.Diagnostics;

namespace InsertIntoAccDB
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static int Main(string[] args)
        {
            OleDbDataAdapter oledbWrite = null;
            OleDbCommand oledbCmd = null;
            int wait = 1000, Num;
            string
                TestID = "", Build = "", ResultColor = "http://ai2s_spps/AI2%20Pictures/TestInProgress.gif",
                Result = "",dateandtime = "",
                //Note = "",
                Field1 = "`ID`",
                Field2 = "`Test ID`",
                Field3 = "`Test Type`",TestType="",
                Field4 = "`Test Name`",TestName="",
                Field5 = "`Build`",
                Field6 = "`ResultColor`",
                Field7 = "`Result`",
                //Field8 = "`Result Details`",
                Field9 = "`Test Started`",
                //Field10 = "`Test Completed`",
                Field11 = "`Qa Owner`",QaOwner="",
                Field12 = "`PE Lead`",PELead="",
                Field13 = "`Framework Data Base Key`",FrameWork = "12345",
                Field14 = "`Box`",Box="",
                Field15 = "`Platform`",Platform="",
                Field16 = "`Product`",Product="",
                Field17 = "`Last run`",
                tablePathname = "d:\\SharePoint\\automatedtestdefinitions.accdb",
                tablename = "`Automated test definitions`",
                connection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tablePathname + ";Persist Security Info=False;";
            
            // begin
 //           Process[] name = Process.GetProcessesByName("InsertIntoAccDB");
   //         if (name.Length > 0) name[0].Kill();
            dateandtime = DateTime.Now.ToString();
            Num = args.Length;
            if (args.Length == 2)
            {
                TestID = args[0];
                Build = args[1];
               // MessageBox.Show("TestID = " + TestID + "; Build = " + Build);
            }
            else
            {
                return 1;
            }

            if (!GetTestInfoFromDB(ref TestID, ref connection, ref tablename, ref Field1, ref TestType, ref TestName, ref Platform, ref QaOwner,
                ref PELead, ref Box, ref Product, ref wait))
            {
                LogToFile("Failed in GetTestInfoFromDB function. "+TestID+"; "+Build);
                return 1;
            }
           // MessageBox.Show("Test Type: " + TestType+"; Test Name: "+TestName+"; QA Owner: "+QaOwner+"; PE Lead: "+PELead+"; Machine: "+Box+"; Platform: "+Platform+"; Product: "+Product);

            if (!SetLastRunInDB(ref TestID, ref tablename, ref Field1, ref oledbCmd, ref oledbWrite, ref dateandtime, ref Field17, ref wait, ref connection, ref Build))
            {
                LogToFile("Failed in SetLastRunInDB function. " + TestID + "; " + Build);
                return 1;
            }
           
            dateandtime = DateTime.Now.ToString();
            if (!InsertRowInDB(ref TestID, ref Build, ref Field3, ref TestType, ref Field4, ref TestName, ref Field5, ref Field6, ref ResultColor, ref Field7,
                ref Result, ref Field9, ref dateandtime, ref Field11, ref QaOwner, ref Field12, ref PELead, ref Field13, ref FrameWork, ref Field14, ref Box,
                ref Field15, ref Platform, ref Field16, ref Product, ref Field2, ref wait))
            {
                LogToFile("Failed in InsertRowInDB. " + TestID + "; " + Build);
                return 1;
            }

            return 0;
        }

        static public bool GetTestInfoFromDB(ref string sTestID, ref string connection, ref string tablename, ref string Field1, ref string TestType, 
            ref string TestName,ref string Platform, ref string QAOwner, ref string PELead, ref string Box, ref string Product, ref int wait)
        {   
            OleDbCommand oledbCmdLocal=null;
            OleDbConnection soledbCnn = null;
            string sql;
            int iTestID;
            bool loop, success = false;
            loop = true;
            iTestID=Convert.ToInt32(sTestID);
            sql = "SELECT * FROM " + tablename + " WHERE " + Field1 + " = " + sTestID;
                     
                try
                {
                    soledbCnn = new OleDbConnection(connection);
                    success = true;
                }
                catch (Exception ex)
                {
                    LogToFile("soledbCnn - Cannot open connection ! " + ex.Message);
                    success = false;
                }
                if (success)
                {
                    try
                    {
                        oledbCmdLocal = new OleDbCommand(sql, soledbCnn);
                        success = true;
                    }
                    catch (Exception ex)
                    {
                        LogToFile("oledbCmdLocal - Cannot open connection ! " + ex.Message);
                        System.Threading.Thread.Sleep(wait);
                        success = false;
                    }
                }
                if (success)
                {
                    try
                    {
                        soledbCnn.Open();
                        success = true;
                    }
                    catch (Exception ex)
                    {
                        LogToFile("oledbCnn.Open() - Cannot open connection ! " + ex.Message);
                        System.Threading.Thread.Sleep(wait);
                        success = false;
                    }
                }
                if (success)
                {
                    while (loop)
                    {
                        try
                        {
                            using (OleDbDataReader oledbReader = oledbCmdLocal.ExecuteReader())
                            {
                                loop = false;
                                while (oledbReader.Read())
                                {
                                   // string ID = oledbReader.GetValue(0).ToString();
                                    TestType = oledbReader.GetString(1).Trim();
                                    TestName = oledbReader.GetString(2).Trim();
                                    Platform = oledbReader.GetString(3).Trim();
                                    QAOwner = oledbReader.GetString(4).Trim();
                                    PELead = oledbReader.GetString(5).Trim();
                                    Box = oledbReader.GetString(12).Trim();
                                    Product = oledbReader.GetString(15).Trim();
                                    success = true;
                                }
                                oledbReader.Close();
                                oledbCmdLocal.Dispose();
                                soledbCnn.Close();
                            }
                        }
                        catch (Exception ex)
                        {
                            LogToFile("oledbReader - Cannot open connection ! " + ex.Message);
                            System.Threading.Thread.Sleep(wait);
                            loop = true;
                            success = false;
                        }
                    }
                }
            oledbCmdLocal.Dispose();
            soledbCnn.Close();
            if (!success) return false;
            //oledbReader.Close();
            System.Threading.Thread.Sleep(wait);
            return true;
        }

        static bool SetLastRunInDB(ref string TestID, ref string tablename, ref string Field1,ref OleDbCommand oledbCmd,
            ref OleDbDataAdapter oledbWrite, ref string dateandtime, ref string Field17, ref int wait,ref string connection, ref string Build)
        {
            string SQL;
            bool loop, success=false;
            OleDbConnection oledbCnn = null;
            SQL = "UPDATE " + tablename + " SET " + Field17 + " = '" + dateandtime + " "+ Build + "' WHERE " + Field1 + " = " + TestID + "";
            try
            {
                oledbCnn = new OleDbConnection(connection);
                success = true;
            }
            catch (Exception ex)
            {
                LogToFile("oledbCnn - Cannot open connection ! " + ex.Message);
                success = false;
            }
            
            loop = true;
            if (success)
            {
                while (loop)
                {
                    try
                    {
                        oledbCmd = new OleDbCommand(SQL, oledbCnn);
                        loop = false;
                        success = true;
                    }
                    catch (Exception ex)
                    {
                        LogToFile("oledbCmd - Cannot open connection ! " + ex.Message);
                        System.Threading.Thread.Sleep(wait);
                        loop = true;
                        success = false;
                    }


                    try
                    {

                        oledbCmd.Connection.Open();
                        oledbCmd.ExecuteNonQuery();
                        loop = false;
                        success = true;

                        oledbCmd.Dispose();
                        oledbCnn.Close();
                    }
                    catch (Exception ex)
                    {
                        LogToFile("oledbDataAdapter - Cannot open connection ! " + ex.Message);
                        System.Threading.Thread.Sleep(wait);
                        loop = true;
                        success = false;
                    }
                }
            }
            oledbCmd.Dispose();
            oledbCnn.Close();
            //if (success) MessageBox.Show(Field2 + " = " + str_str);
            if (!success) return false;
            return true;
        }

        static bool InsertRowInDB(ref string TestID, ref string Build, ref string Field3, ref string TestType, ref string Field4, ref string TestName, 
            ref string Field5, ref string Field6, ref string ResultColor, ref string Field7, ref string Result, ref string Field9, ref string dateandtime,
            ref string Field11, ref string QaOwner, ref string Field12, ref string PELead, ref string Field13, ref string FrameWork, ref string Field14,
            ref string Box, ref string Field15, ref string Platform, ref string Field16, ref string Product, ref string Field2, ref int wait)
        {
            string connection, tablePathname, tablename,SQL;
            bool loop, success = false;
            OleDbConnection oledbCnn = null;
           // OleDbDataReader oledbReader = null;
           // OleDbDataAdapter oledbWrite = null;
            OleDbCommand oledbCmd = null;

            if (TestType.Contains("Regression"))
            {
                tablePathname = "D:\\SharePoint\\TestResults\\11Regression.accdb";
                tablename = "`TestAutomationResults-11Regression`";
            }
            else
            {
                tablePathname = "D:\\SharePoint\\TestResults\\11Other.accdb";
                tablename = "`TestAutomationResults-11Other`";
            }

            connection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tablePathname + ";Persist Security Info=False;";
            SQL="INSERT INTO "+tablename+" ("+Field2+", "+Field3+", "+Field4+", "+Field5+", "+Field6+", "+Field7+", "+Field9+", "+Field11+", "+
                Field12+", "+Field13+", "+Field14+", "+Field15+", "+Field16+") VALUES ('"+TestID+"', '"+TestType+"', '"+TestName+"', '"+Build+
                "', '" +ResultColor+ "', 'In Process', '" + dateandtime + "', '"+QaOwner+"', '"+PELead+"', '"+FrameWork+"', '"+Box+"', '"+Platform+"', '"+Product+"')";

            success = false;

            loop = true;
            while (loop)
            {
                try
                {
                    oledbCnn = new OleDbConnection(connection);
                    success = true;
                }
                catch (Exception ex)
                {
                    LogToFile("oledbCnn - Cannot open connection ! " + ex.Message);
                    System.Threading.Thread.Sleep(wait);
                    loop = true;
                    success = false;
                }

                if (success)
                {
                    try
                    {
                        oledbCmd = new OleDbCommand(SQL, oledbCnn);
                        loop = false;
                        success = true;
                    }
                    catch (Exception ex)
                    {
                        LogToFile("oledbCmd - Cannot open connection ! " + ex.Message);
                        System.Threading.Thread.Sleep(wait);
                        loop = true;
                        success = false;
                    }
                }
                if (success)
                {
                    try
                    {

                        oledbCmd.Connection.Open();
                        oledbCmd.ExecuteNonQuery();
                        loop = false;
                        success = true;

                        oledbCmd.Dispose();
                        oledbCnn.Close();
                    }
                    catch (Exception ex)
                    {
                        LogToFile("oledbDataAdapter - Cannot open connection ! " + ex.Message);
                        System.Threading.Thread.Sleep(wait);
                        loop = true;
                        success = false;
                    }
                }
            }

            oledbCmd.Dispose();
            oledbCnn.Close();
            if (!success) return false;
            
            return true;

        }

        static public void LogToFile(string LogThis)
        {
            string FileNamePath = "D:\\ShareThis\\InsertIntoACCLog.txt";
            using (StreamWriter fileStream = File.AppendText(FileNamePath))
            {
                fileStream.WriteLine("{0} {1}", DateTime.Now.ToString(), LogThis);
                fileStream.Flush();
                fileStream.Close();
            }
        }
    }
}