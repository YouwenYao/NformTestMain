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
		
		/// </summary>
		
		/// <param name="sourceDir">source directory</param>
		
		/// <param name="targetDir">target directory</param>
		
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
        /// Backup database.
        /// </summary>
    	private void BackupDB (String RestoreDB, String ServerLogPath, String ViewerLogPath)
		{
			// If RestoreDB is Y, program will restore Database for Nform before scripts are executed.
           if(RestoreDB.Equals("Y"))
           {
           	 	//stop Nform service
				Console.WriteLine("Stop Nform service...");
				string strRst = RunCommand("sc stop Nform");
				//First, delete log files to clean the Nform.   
                Delay.Duration(10000);
                
                //If NformService is still running, stop Nform again.
                if(!isStopped())
                {
                	RunCommand("sc stop Nform");
                	Delay.Duration(10000);
                }
				
                myLxDBOper.DeleteLogFile(ServerLogPath);
                myLxDBOper.DeleteLogFile(ViewerLogPath);
				
				//Backup Database operation. Just do once before run all scripts.
	            myLxDBOper.SetDbType();
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
				
				//If Nform is not running, start Nform again.
				if(!isStarted())
				{
					//start Nform service
		            Console.WriteLine("Start Nform service...");
					RunCommand("sc start Nform");		
				}
           }
		}
    	
    	/// <summary>
        /// Restore database.
        /// </summary>
    	private void RestoreDB (String RestoreDB, String ServerLogPath, String ViewerLogPath)
		{			
    		// If RestoreDB is Y, program will restore Database for Nform before scripts are executed.
           if(RestoreDB.Equals("Y"))
           {
	            //stop Nform service
				Console.WriteLine("Stop Nform service...");
				RunCommand("sc stop Nform");
				Delay.Duration(10000);
				
				//If NformService is still running, stop Nform again.
                if(!isStopped())
                {
                	RunCommand("sc stop Nform");
                	Delay.Duration(10000);
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
				
				//If Nform is not running, start Nform again.
				if(!isStarted())
				{
					//start Nform service
		            Console.WriteLine("Start Nform service...");
					RunCommand("sc start Nform");		
				}				
           }
		}
    	
    	
    	//**********************************************************************
		/// <summary>
		/// Check Nform process is stopped.
		/// </summary>
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
            string RestoreDB_Flag = configs["RestoreDB_AfterEachTestCase"];
            string ServerLogPath = configs["Server_Log_Path"];
            string ViewerLogPath = configs["Viewer_Log_Path"];
            string DetailSteps = configs["DetailSteps_InResult"];
            BackupDB(RestoreDB_Flag, ServerLogPath,ViewerLogPath);
            
            string tsName = mainOp.getTestCaseName();
            string excelPath = "keywordscripts/" + tsName + ".xlsx"; 
            string keywordName = "keywordscripts/";
            Report.Info("INfo",excelPath);	
            
            mainOp.StrExcelDirve = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(),
                                                 excelPath);
            mainOp.runApp();							//  ********* 1. run Application *********
            
            //Create Report folder
            string reportDir = System.IO.Directory.GetCurrentDirectory();
            System.IO.DirectoryInfo reportDirect = System.IO.Directory.CreateDirectory(reportDir + @"\" +"Report_" + System.DateTime.Now.ToString ("yyyyMMdd_HHmmss")); 
            string reportPath = reportDirect.FullName+@"\";
            string sourcePath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(),
                                                 keywordName);
            LxParse stepsRepository = mainOp.getSteps();
            // stepsRepository.doValidate();  				//  ********* 2. check scripts  syntax *********
           
            ArrayList stepList = stepsRepository.getStepsList();
            bool result = LxGenericAction.performScripts(stepList,DetailSteps);	//  ********* 3. run scripts with reflection *********
           
            mainOp.setResult();
            mainOp.runOverOneCase(tsName);
            mainOp.opXls.close();
            Delay.Seconds(5);
            LxTearDown.closeApp(mainOp.ProcessId);		//  ********* 4. clean up for next running *********
			
            RestoreDB(RestoreDB_Flag, ServerLogPath, ViewerLogPath);
           
            // Add excel file link in Ranorex.report.
            string linkFile = reportPath + tsName +  ".xlsx";
			string html = String.Format("<a href='{0}'>{1}</a>",linkFile,tsName);
			Report.LogHtml(ReportLevel.Info, "Info", html);
            
            //Copy all report created by Ranorex to Nform Report folder.
            XCopy(sourcePath, reportPath);
            
            
            if(!result)
            {
				throw new Ranorex.ValidationException("The test case running failed!", null);  
            }
        }
        
    }
}
