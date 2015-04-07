/*
 * Created by Ranorex
 * User: y93248
 * Date: 2011-11-15
 * Time: 15:07
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */

using System;
using System.Threading;
using System.Drawing;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using WinForms = System.Windows.Forms;
using Emerson.NetworkPower.Velocity;
using Lextm.SharpSnmpLib;
using Lextm.SharpSnmpLib.Messaging;
using System.Net;
using System.Net.Sockets;
using System.Windows.Forms;
using System.Diagnostics;
using System.Configuration;


using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Reporting;
using Ranorex.Core.Testing;
using NformTester.lib;
using System.Data.SqlClient;

namespace NformTester
{	
/// <summary>
/// Main in Program Class
/// </summary>
   class Program
    {
    	/// <summary>
        /// Get all info from app.config
        /// </summary>
    	private static IDictionary<string, string> GetConfigs ()
		{
			var configs = new Dictionary<string, string> ();
			int len = ConfigurationManager.AppSettings.Count;
			for (int i = 0; i < len; i++)
			{
				configs.Add (
					ConfigurationManager.AppSettings.GetKey (i),
					ConfigurationManager.AppSettings[i]);
			}

			return configs;
		}
    	
    	//Report name
    	/// <summary>
    	/// Report Folder.
    	/// </summary>
    	public static string ProReportFolder = "";
    	
    	
    	
    	/// <summary>
        /// Get Report folder
        /// </summary>
        /// <returns>Report Folder</returns>
    	public static string getReport()
    	{
    		return ProReportFolder;
    	}
    	
    	/// <summary>
    	/// Set Report folder
    	/// </summary>
    	/// <param name="path">Report Path</param>
    	public static void setReport(string path)
    	{
    		ProReportFolder = path; 
    	}
    	
    	
    	/// <summary>
        /// Main program
        /// </summary>
        /// <returns>Error Code</returns>
        /// <param name="args">args</param>
    	[STAThread]
        public static int Main(string[] args)
        {
 			/*
        	if(!CheckDeviceAvailable())
        	{
        		DialogResult dr = MessageBox.Show("Do you want to continue? Click Ok button to continue.", "Warning" , MessageBoxButtons.YesNo,MessageBoxIcon.Warning,MessageBoxDefaultButton.Button2);
            	if(dr == DialogResult.No)
            	{
            		return 0;
            	}
        	}                   
            
             var configs = GetConfigs ();
             string CheckDevice = configs["CheckDevice_BeforeTesting"];
             string RestoreDB = configs["RestoreDB_AfterEachTestCase"];
             */
            
            /*
             * Get size of database for SQL SERVER
     		LxSQLDbOperation SQLOper = new LxSQLDbOperation();
     		SqlConnection conn = new SqlConnection();
	        conn.ConnectionString =@"Data Source=10.146.64.56\SQLEXPRESS;Initial Catalog=master;User ID=sa;Password=sa@2013;";
			SQLOper.OpenConnection(conn);
     		string NformSize = SQLOper.GetDbSize(conn, "Nform");
     		string NformAlmSize = SQLOper.GetDbSize(conn, "NformAlm");
     		string NformLogSize = SQLOper.GetDbSize(conn, "NformLog");
     		Console.WriteLine("The size of Nform is:"+NformSize);
     		Console.WriteLine("The size of Nform is:"+NformAlmSize);
     		Console.WriteLine("The size of Nform is:"+NformLogSize);
     		*/
     		
     		/*
     		 * Get size of database for SQL CE
     		LxCEDbOperation CEOper = new LxCEDbOperation();
     		double NformAlmSize = CEOper.GetAlarmDbSize();
     		double NformDataLogSize = CEOper.GetDataLogDbSize();
			Console.WriteLine("NformAlmSize is: " +NformAlmSize);
     		Console.WriteLine("NformDataLogSize is: " +NformDataLogSize);
     		 * */

     	   /*
     	    * Increase database for SQLCE 
     	    LxCEDbOperation CEOper = new LxCEDbOperation();
			CEOper.IncreaseAlarmDbSize();
			CEOper.IncreaseDatalogDbSize();
     	    * */
     	   
     	   /*
     	    * Increase database for SQL SERVER 
     	    LxSQLDbOperation SQLOper = new LxSQLDbOperation();
			SqlConnection conn = new SqlConnection();
	        conn.ConnectionString =@"Data Source=10.146.64.56\SQLEXPRESS;Initial Catalog=master;User ID=sa;Password=sa@2013;";
			SQLOper.GetDbSize(conn,"NformAlm");
			SQLOper.GetDbSize(conn,"NformLog");
			SQLOper.IncreaseAlarmDbSize();
			SQLOper.IncreaseDatalogDbSize();
     	    * 
     	    * */
     	   
     	   /*
     	    * Get table value for SQLCE database
     	    *      	   
     	    LxCEDbOperation CEOper = new LxCEDbOperation();
     	    string dbNformName = @"Nform.sdf";
     	    string dbNformAlmName = @"NformAlm.sdf";
     	    string dbNformLogName = @"NformLog.sdf";
     	    string cmdVersion = @"SELECT * FROM Version;";
     	    string cmdAlarm = @"SELECT * FROM Alarm;";
     	    string cmdLog = @"SELECT * FROM DataLog;";
	   	    CEOper.GetTableValue(dbNformName,cmdVersion);
 	  	    CEOper.GetTableValue(dbNformAlmName,cmdAlarm);
    	    CEOper.GetTableValue(dbNformLogName,cmdLog);
     	    * 
     	    * */
     	   
     	   /*
     	    * Get table value for SQLServer database
     	    LxSQLDbOperation SQLOper = new LxSQLDbOperation();
			SqlConnection conn = new SqlConnection();
	        conn.ConnectionString =@"Data Source=10.146.64.56\SQLEXPRESS;Initial Catalog=master;User ID=sa;Password=sa@2013;";
	        string cmdVersion = @"use Nform;SELECT * FROM Version;";
	        string cmdAlarm = @"use NformAlm;SELECT * FROM Alarm;";
	        string cmdLog = @"use NformLog;SELECT * FROM DataLog;";
	        string cmdGrp = @"use Nform;SELECT * FROM UsrGrp;";  
	        SQLOper.GetTableValue(conn,cmdVersion);
	        SQLOper.GetTableValue(conn,cmdGrp);
     	    * */  
     	/*   
     	   string ip1 = AppConfigOper.parseToValue("$SNMP_SingleAuto_1$");
     	   string ip2 = AppConfigOper.parseToValue("$Velocity_device_2$");
     	   Console.WriteLine("ip1 is:"+ip1);
     	   Console.WriteLine("ip2 is:"+ip2);
     	*/   
     		string CheckDevice = AppConfigOper.mainOp.getConfigValue("CheckDevice_BeforeTesting");
            string RestoreDB = AppConfigOper.mainOp.getConfigValue("RestoreDB_AfterEachTestCase");
            string runOnVM = AppConfigOper.mainOp.getConfigValue("RunOnVM");
             //Create Report folder
            string reportDir = System.IO.Directory.GetCurrentDirectory();
            System.IO.DirectoryInfo reportDirect = System.IO.Directory.CreateDirectory(reportDir + @"\Report\" +"Report_" + System.DateTime.Now.ToString ("yyyyMMdd_HHmmss")); 
     	    string ReportPath = reportDirect.FullName+@"\"; 
     		setReport(ReportPath);
     
            // If CheckDevice is Y, program will check these ip addresses are available or not.
             if(CheckDevice.Equals("Y"))
             {
	           //stop Nform service
				Console.WriteLine("Stop Nform service...");
				string strRst = RunCommand("sc stop Nform");
			   //Be used to check devices are avalibale or not, which are configured in Device.ini
	           LxDeviceAvailable myDeviceAvailable = new LxDeviceAvailable();
	       	   myDeviceAvailable.CheckSnmpDevice();
	           myDeviceAvailable.CheckVelDevice();
	           //start Nform service
	           Console.WriteLine("Start Nform service...");
			   strRst = RunCommand("sc start Nform");	
             }                        
         
             if(runOnVM.Equals("Y"))
             {
             	Keyboard.Enabled = false;  
             	Mouse.Enabled = false;
             	Keyboard.AbortKey = System.Windows.Forms.Keys.Pause;  
             	NformRepository.Instance.SearchTimeout = new Duration(50000);
             }
             
        	Keyboard.AbortKey = System.Windows.Forms.Keys.Pause;
            int error = 0;
            /*
            try
            {
                error = TestSuiteRunner.Run(typeof(Program), Environment.CommandLine);            	
            }
            catch (Exception e)
            {
               MessageBox.Show("Unexpected exception occurred:");
            	Report.Error("Unexpected exception occurred: " + e.ToString());
                error = -1;
            }
            */
           error = TestSuiteRunner.Run(typeof(Program), Environment.CommandLine);
           
            return error;
        }
        
         //**********************************************************************
		/// <summary>
		/// Run cmd command
		/// </summary>
		/// <returns>Result</returns>
		public static string RunCommand(string command)
		{
			Process p = new Process();
			p.StartInfo.FileName = "cmd.exe";
			p.StartInfo.UseShellExecute = false;
			p.StartInfo.RedirectStandardInput = true;
			p.StartInfo.RedirectStandardOutput = false;
			p.StartInfo.RedirectStandardError = true;
			p.StartInfo.CreateNoWindow = true;
			p.Start();
			p.StandardInput.WriteLine(command);
			p.StandardInput.WriteLine("exit");
			Delay.Duration(10000);
			return  "";
		}
        
    }
}
