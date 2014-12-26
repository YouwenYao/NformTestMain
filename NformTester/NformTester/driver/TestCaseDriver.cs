/*
 * Created by Ranorex
 * User: y93248
 * Date: 2012-3-12
 * Time: 13:51
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using System.Collections;
using WinForms = System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Win32;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;
using NformTester.lib;
using System.Windows.Forms;

namespace NformTester.driver
{
	
	
    /// <summary>
    /// Description of TestCaseDriver.
    /// </summary>TST1059TST1059
    [TestModule("78CE2FBF-D204-4698-98BF-68118D374E9F", ModuleType.UserCode, 1)]
    public class TestCaseDriver : ITestModule
    {
    	/// <summary>
		/// Script path
		/// </summary>
    	public string excelPath = "";
    	
    	/// <summary>
        /// Result of backup database. Used by restore database of every script.
        /// </summary>
    	public LxDBOper myLxDBOper = new LxDBOper();
    	
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public TestCaseDriver()
        {
            // Do not delete - a parameterless constructor is required!
        }
        
		
		/// for an example：XCopy("c:\a\", "d:\b\");
		/// <summary>Copy source script to report folder</summary>
		/// <param name="sourceDir">sourceDir</param>
		/// <param name="targetDir">targetDir</param>
		public static void XCopy(string sourceDir, string targetDir)
		
    	{
		   //If the source directory exists.
		   if (System.IO.Directory.Exists(sourceDir))
		   {
		       //If the source directory does not exist, create it.
		       if (!System.IO.Directory.Exists(targetDir))
		       System.IO.Directory.CreateDirectory(targetDir);
		       //Get data from sourcedir.
		       System.IO.DirectoryInfo sourceInfo = new System.IO.DirectoryInfo(sourceDir);
		       //Copy the files.
		       System.IO.FileInfo[] files = sourceInfo.GetFiles();
		       foreach (System.IO.FileInfo file in files)
		       {
		           System.IO.File.Copy(sourceDir + "\\" + file.Name, targetDir + "\\" + file.Name, true);
		       }
		       //Copy the dir.
		       System.IO.DirectoryInfo[] dirs = sourceInfo.GetDirectories();
		       foreach (System.IO.DirectoryInfo dir in dirs)
		       {
		          string currentSource = dir.FullName;
		          string currentTarget = dir.FullName.Replace(sourceDir, targetDir);
		          System.IO.Directory.CreateDirectory(currentTarget);
		          //recursion
		          XCopy(currentSource, currentTarget);
		        }
		      }
		   }
		

		/// <summary>
		/// Backup database
		/// </summary>
    	private void BackupDB ()
		{
    		string RestoreDB = AppConfigOper.getConfigValue("RestoreDB_AfterEachTestCase");
            string ServerLogPath = AppConfigOper.getConfigValue("Server_Log_Path");
            string ViewerLogPath = AppConfigOper.getConfigValue("Viewer_Log_Path");
            string DB_DbType = AppConfigOper.getConfigValue("DB_DbType");
    		// If RestoreDB is Y, program will restore Database for Nform before scripts are executed.
           if(RestoreDB.Equals("Y"))
           {
           	 	//stop Nform service
				Console.WriteLine("Stop Nform service...");
				string strRst = RunCommand("sc stop Nform");
				//First, delete log files to clean the Nform.   
                Delay.Duration(10000);
                
                //If NformService is still running, stop Nform again. Try 3 times.
                for(int i=0;i<3;i++)
                {
                	if(!isStopped())
                	{
	                	RunCommand("sc stop Nform");
	                	Delay.Duration(10000);
               		}
                }
				
                myLxDBOper.DeleteLogFile(ServerLogPath);
                myLxDBOper.DeleteLogFile(ViewerLogPath);
				
				//Backup Database operation. Just do once before run all scripts.
	            myLxDBOper.SetDbType(DB_DbType);
	            myLxDBOper.BackUpDataBase();	
	           if(myLxDBOper.GetBackUpResult() == false)
	            {
	               Console.WriteLine("Back up database is faild!");
	            }
	            else
	            {
	            	Console.WriteLine("Back up database is successful!");
	            }
	            //start Nform service
	            Console.WriteLine("Start Nform service...");
				strRst = RunCommand("sc start Nform");	
				
				//If Nform is not running, start Nform again. Try 3 times.
				for(int j=0;j<3;j++)
				{
					if(!isStarted())
					{
						//start Nform service
			            Console.WriteLine("Start Nform service...");
						RunCommand("sc start Nform");		
					}
					
				}
           }
		}

    	/// <summary>
    	/// Restore Database
    	/// </summary>
    	private void RestoreDB ()
		{			
    		string RestoreDB = AppConfigOper.getConfigValue("RestoreDB_AfterEachTestCase");
    		string ServerLogPath = AppConfigOper.getConfigValue("Server_Log_Path");
    		string ViewerLogPath = AppConfigOper.getConfigValue("Viewer_Log_Path");
    		// If RestoreDB is Y, program will restore Database for Nform before scripts are executed.
           if(RestoreDB.Equals("Y"))
           {
	            //stop Nform service
				Console.WriteLine("Stop Nform service...");
				RunCommand("sc stop Nform");
				Delay.Duration(10000);
                
                 //If NformService is still running, stop Nform again. Try 3 times.
                for(int i=0;i<3;i++)
                {
                	if(!isStopped())
                	{
	                	RunCommand("sc stop Nform");
	                	Delay.Duration(10000);
               		}
                }
				
				
                //First, delete log files to clean the Nform.             
      //          myLxDBOper.DeleteLogFile(ServerLogPath);
      //          myLxDBOper.DeleteLogFile(ViewerLogPath);
                
	            // Restore Database operation.
	            // If there is any error when perform Scripts, execute the restore DB operation.
	            myLxDBOper.RestoreDataBase();
	            if(myLxDBOper.GetRestoreResult() == false)
	            {
	               Console.WriteLine("Restore database is faild! You need to restore database manually");
	            }
	            else
	            {
	            	Console.WriteLine("Restore database is successful!");
	            }
	            //start Nform service
	            Console.WriteLine("Start Nform service...");
				RunCommand("sc start Nform");	
				Delay.Duration(10000);
				
				//If Nform is not running, start Nform again. Try 3 times.
				for(int j=0;j<3;j++)
				{
					if(!isStarted())
					{
						//start Nform service
			            Console.WriteLine("Start Nform service...");
						RunCommand("sc start Nform");		
					}
					
				}
           }
		}
    	
    	
    	//**********************************************************************
		/// <summary>
		/// Check Nform process is stopped.
		/// </summary>
		/// <returns>isStopped</returns>
		public static bool isStopped()
		{
    	        Process [] pList = Process.GetProcesses(); 
				bool isStopped = true; 
				foreach(Process p in pList) 
				{ 
					if(p.ProcessName == "NformServer") 
					{ 
						isStopped = false;
						break;
					} 
				} 
				return isStopped;
		}
		
		
		//**********************************************************************
		/// <summary>
		/// Check Nform process is started.
		/// </summary>
		/// <returns>isStarted</returns>
		public static bool isStarted()
		{
    	        Process [] pList = Process.GetProcesses(); 
				bool isStarted = false; 
				foreach(Process p in pList) 
				{ 
					if(p.ProcessName == "NformServer") 
					{ 
						isStarted = true;
						break;
					} 
				} 
				return isStarted;
		}
		
    	
    	//**********************************************************************
		/// <summary>
		/// Run cmd command
		/// </summary>
		/// <param name="command">command</param>
		/// <returns>command</returns>
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
		
        /// <summary>
        /// Performs the playback of actions in this module.
        /// </summary>
        /// <remarks>You should not call this method directly, instead pass the module
        /// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
        /// that will in turn invoke this method.</remarks>
        void ITestModule.Run()
        {           
        	Mouse.DefaultMoveTime = 300;
            Keyboard.DefaultKeyPressTime = 100;
            Delay.SpeedFactor = 1.0;                                 
            LxSetup mainOp = LxSetup.getInstance();  
            var configs = mainOp.configs;
            
            string DetailSteps = AppConfigOper.getConfigValue("DetailSteps_InResult");
            string strUserName =  AppConfigOper.getConfigValue("UserName");
            string strPassword =  AppConfigOper.getConfigValue("Password");
            string strServerName =  AppConfigOper.getConfigValue("ServerName");
            
            BackupDB();
            
            string tsName = mainOp.getTestCaseName();
            string excelPath = "keywordscripts/" + tsName + ".xlsx"; 
            string keywordName = "keywordscripts/";
            Report.Info("INfo",excelPath);	
            
            mainOp.StrExcelDirve = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(),
                                                 excelPath);
            mainOp.m_strUserName = strUserName;
            mainOp.m_stPassword = strPassword;
            mainOp.m_strServerName = strServerName;
            
            mainOp.runApp();//  ********* 1. run Application *********
            
            //Copy script to Report folder
            string reportPath = Program.getReport();
            string sourcePath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(),
                                                 keywordName);
            
            string reprotFileName = reportPath + tsName + ".xlsx";
            string sourceFileName = mainOp.StrExcelDirve;
            LxParse stepsRepository = mainOp.getSteps();
            // stepsRepository.doValidate();  				//  ********* 2. check scripts  syntax *********
           
            int myProcess = mainOp.ProcessId;
            ArrayList stepList = stepsRepository.getStepsList();
            bool result = true;	
            //  ********* 3. run scripts with reflection *********
 			try
 			{
 				result = LxGenericAction.performScripts(stepList,DetailSteps);
 			}
 			catch(Exception e)
 			{
 				result = false;
 				LxLog.Error("Error",tsName+" "+e.Message.ToString());
 			}

            mainOp.setResult();
            mainOp.runOverOneCase(tsName);
            mainOp.opXls.close();
            Delay.Seconds(5);
            LxTearDown.closeApp(mainOp.ProcessId);		//  ********* 4. clean up for next running *********
			
            RestoreDB();
           
            // Add excel file link in Ranorex.report.
            string linkFile = reportPath + tsName +  ".xlsx";
			string html = String.Format("<a href='{0}'>{1}</a>",linkFile,tsName);
			Report.LogHtml(ReportLevel.Info, "Info", html);
            
            //Copy all report created by Ranorex to Nform Report folder.
//          XCopy(sourcePath, reportPath);
            System.IO.File.Copy(sourceFileName,reprotFileName,true);
            
            if(!result)
            {
				throw new Ranorex.ValidationException("The test case running failed!", null);  
            }
        }
        
    }
}
