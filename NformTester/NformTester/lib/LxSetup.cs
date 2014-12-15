/*
 * Created by Ranorex
 * User: y93248
 * Date: 2011-11-16
 * Time: 13:47
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;
using System.Xml;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Configuration;
using System.Windows.Forms;
using System.Xml;

namespace NformTester.lib
{
	//**************************************************************************
	/// <summary>
	/// Setup the environment for scripts running
	/// </summary>
	/// <para> Author: Peter Yao</para>
	public class LxSetup
	{
		/// <summary>
		/// The only one instance of this class
		/// </summary>
		public static LxSetup m_instance;
		
		/// <summary>
		/// The test object Login username
		/// </summary>
		public string m_strUserName = "";
		
		/// <summary>
		/// The test object Login password
		/// </summary>
		public string m_stPassword = "";
		
		/// <summary>
		/// The test object Login server name
		/// </summary>
		public string m_strServerName = "";
		
		
		/// <summary>
		/// The test object application file path
		/// </summary>
		public string m_strApplicationName = "";
		
		/// <summary>
		/// The path of scripts file
		/// </summary>
		public string m_strExcelDirve = "";
		
		/// <summary>
		/// The process id of application
		/// </summary>
		public int m_iProcessId = 0;
		
		/// <summary>
		/// The instance to operate Xls
		/// </summary>
		public LxXlsOper opXls = null;
		
		/// <summary>
		/// The testcase runlist from rxtst file
		/// </summary>
		public ArrayList m_Runlist = new ArrayList();
		
		/// <summary>
		/// Numbers of commands
		/// </summary>
		public int m_iRowNum = 0;
		
		/// <summary>
		/// Start line of commands
		/// </summary>
		public int m_iRowStart = 0;
		
		/// <summary>
		/// End line of commands
		/// </summary>
		public int m_iRowEnd = 0;
		
		/// <summary>
		/// Pre-condition.
		/// </summary>
		public string m_strPrecondition = "";
		
		
		/// <summary>
		/// Configs in app.config
		/// </summary>
		public  IDictionary<string, string> configs ;
		
		/// <summary>
		///  Get method for ProcessId
		/// </summary>
		public int ProcessId {
			get { return m_iProcessId; }
		}
		
		/// <summary>
		/// Get and Set method for ApplicationName
		/// </summary>
		public string StrApplicationName {
			get { return m_strApplicationName; }
			set { m_strApplicationName = value; }
		}

		/// <summary>
		/// Get and Set method for ExcelDirve
		/// </summary>
		public string StrExcelDirve {
			get { return m_strExcelDirve; }
			set { m_strExcelDirve = value; }
		}
		
		//**********************************************************************
		/// <summary>
		/// Private Constructer, get the testcase runlist from rxtst file
		/// </summary>
		private LxSetup()
		{
			m_Runlist.Clear();
			m_Runlist = getRunlist();
			configs = GetConfigs();
		}

		/// <summary>
		/// Return the only one instance of this class
		/// </summary>
		/// <returns>m_instance</returns>
		public static LxSetup getInstance() 
		{
			if ( m_instance == null )
			{
				m_instance = new LxSetup();
			} // end if
			return m_instance;
		} // end getInstance

		/// <summary>
		/// Replace the name with value refer to app.config
		/// </summary>
		/// <param name="name">name</param>
		/// <returns>string</returns>
		private string parseToValue(string name)
        {
			LxSetup mainOp = LxSetup.getInstance();
			var configs = mainOp.configs;
			
            string addr = name;
            if (name.Substring(0, 1) == "$" && name.Substring(name.Length - 1, 1) == "$")
            {
                string key = name.Substring(1, name.Length - 2);
                string result = null;
                if(configs.ContainsKey(key))
                {
                	result = configs[key];
                }
                else
                {
                	result = configs["Default"];
                }
           		addr = result;
            }

            return addr.Replace("\"","");
		}

		/// <summary>
		/// Read the application path from scripts and run
		/// </summary>
		public void runApp()
		{
			if(!File.Exists(m_strExcelDirve))
			{
				runOverOneCase(getTestCaseName());
				throw new Ranorex.ElementNotFoundException("The file " + m_strExcelDirve + " not found", null);   
			}
			opXls = new LxXlsOper();
        	opXls.open(m_strExcelDirve);
        	m_strApplicationName = opXls.readCell(2,2);
        	Console.WriteLine(m_strApplicationName);

			m_iRowNum = Convert.ToInt16(opXls.readCell(7,2));
			
			m_strPrecondition = opXls.readCell(23,2);
			
			
			string runningRange = opXls.readCell(8,2);
			Regex rex = new Regex("[0-9]+");
			
			
			if(runningRange == "" || runningRange.IndexOf("a") != -1)
			{
				m_iRowStart = 1;
				m_iRowEnd = m_iRowNum;
			}
			else{
				string[] range = runningRange.Split('-');
				m_iRowStart = Convert.ToInt16(range[0]);
				if(runningRange.IndexOf("-") == -1)
				{
					m_iRowEnd = m_iRowNum;
				}
				else
				{
					m_iRowEnd = Convert.ToInt16(range[1]);
				}
			}
		
			if(m_iRowStart == 1 && m_strApplicationName!="")
			{
				m_iProcessId = Host.Local.RunApplication(parseToValue(m_strApplicationName));
			}
			else
			{
				m_iProcessId = 0;
			}
			
			//Excute Precondition script to verify the Precondtion.
			if (!(m_strPrecondition.Trim().Equals("")))
				VerifyCondition();
			
		}
		

		/// <summary>
		/// Get all info from app.config.
		/// </summary>
		/// <returns>configs</returns>
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
    	
		//**********************************************************************
		/// <summary>
		/// Read the command from scripts
		/// </summary>
		public LxParse getSteps()
		{
			LxParse parse = new LxParse();
        	for( int i = m_iRowStart; i <= m_iRowEnd; i++)
        	{
        		LxScriptItem item = new LxScriptItem();
//        		item.m_Index = opXls.readCell(i+1,3).Trim();
//        		item.m_Type = opXls.readCell(i+1,4).Trim();
//        		item.m_WindowName = opXls.readCell(i+1,5).Trim();
//        		item.m_Component = opXls.readCell(i+1,6).Trim();
//        		item.m_Action = opXls.readCell(i+1,7).Trim();
//        		item.m_Arg1 = opXls.readCell(i+1,8).Trim();
//        		item.m_Arg2 = opXls.readCell(i+1,9).Trim();
//        		item.m_Arg3 = opXls.readCell(i+1,10).Trim();
//        		item.m_Arg4 = opXls.readCell(i+1,11).Trim();
//        		item.m_Arg5 = opXls.readCell(i+1,12).Trim();
//        		item.m_Arg6 = opXls.readCell(i+1,13).Trim();
				item.m_Index = opXls.readCell(i+1,3);
        		item.m_Type = opXls.readCell(i+1,4);
        		item.m_WindowName = opXls.readCell(i+1,5);
        		item.m_Component = opXls.readCell(i+1,6);
        		item.m_Action = opXls.readCell(i+1,7);
        		item.m_Arg1 = opXls.readCell(i+1,8);
        		item.m_Arg2 = opXls.readCell(i+1,9);
        		item.m_Arg3 = opXls.readCell(i+1,10);
        		item.m_Arg4 = opXls.readCell(i+1,11);
        		item.m_Arg5 = opXls.readCell(i+1,12);
        		item.m_Arg6 = opXls.readCell(i+1,13);
        		parse.m_Scripts.Add(item);        		
        	}		
			return parse;
		}
		
		//**********************************************************************
		/// <summary>
		/// After one case running, set the result(pass or fail) to the scripts file
		/// </summary>
		public void setResult()
		{
			int i = 1;
			for( i = m_iRowStart; i <= m_iRowEnd; i++)
        	{
        		string commandResult = opXls.readCell(i+1,14).Trim();
        		if(commandResult == "Fail")
        		{
        			break;
        		}	
        	}		
			// writeInfo();
			opXls.writeCell(9,2,i <= m_iRowEnd?"Fail":"Pass");
		}
		
		//**********************************************************************
		/// <summary>
		/// After one case running, write additional info to the scripts file
		/// </summary>
		public void writeInfo()
		{			
			opXls.writeCell(12,2,Ranorex.Host.Local.RanorexVersion);
			opXls.writeCell(21,2,Ranorex.Host.Local.OSEdition + "  " + Ranorex.Host.Local.OSVersion);
			opXls.writeCell(22,2,Ranorex.Host.Local.RuntimeVersion);
			
			
			NformRepository repo = NformRepository.Instance;
			
			//Delay.Milliseconds(2000);
			if(!repo.NFormApp.NformG2Window.FormMain.HelpInfo.Exists())
				return;
			//Delay.Milliseconds(2000);
						
			repo.NFormApp.NformG2Window.FormMain.Help.Click();
			repo.NFormApp.NformG2Window.FormMain.About_Liebert_Nform.Click();
			Ranorex.NativeWindow nativeWnd = repo.NFormApp.Help.FormAbout_LiebertR_Nform.ViewerVerInfo.CreateAdapter<Ranorex.NativeWindow>(false);
			string viewVer = nativeWnd.WindowText;
			repo.NFormApp.Help.FormAbout_LiebertR_Nform.TabServer.Click();		
			nativeWnd = repo.NFormApp.Help.FormAbout_LiebertR_Nform.ServerVerInfo.CreateAdapter<Ranorex.NativeWindow>(false);
			string severVer = nativeWnd.WindowText;
			repo.NFormApp.Help.FormAbout_LiebertR_Nform.TabDatabase.Click();
			nativeWnd = repo.NFormApp.Help.FormAbout_LiebertR_Nform.DbVersionInfo.CreateAdapter<Ranorex.NativeWindow>(false);
			string dbVer = nativeWnd.WindowText;
			nativeWnd = repo.NFormApp.Help.FormAbout_LiebertR_Nform.DbEditionInfo.CreateAdapter<Ranorex.NativeWindow>(false);
			string dbEdition = nativeWnd.WindowText;
			repo.NFormApp.Help.FormAbout_LiebertR_Nform.TabLicense.Click();
			// nativeWnd = repo.NFormApp.Help.FormAbout_LiebertR_Nform.LicenseDetailInfo.CreateAdapter<Ranorex.NativeWindow>(false);
			string licenseDetail = repo.NFormApp.Help.FormAbout_LiebertR_Nform.LicenseDetail.TextValue;//nativeWnd.WindowText;
			repo.NFormApp.Help.FormAbout_LiebertR_Nform.TabRegistration.Click();
			nativeWnd = repo.NFormApp.Help.FormAbout_LiebertR_Nform.RegistInfoDscrInfo.CreateAdapter<Ranorex.NativeWindow>(false);
			string regist = nativeWnd.WindowText;
			repo.NFormApp.Help.FormAbout_LiebertR_Nform.OK.Click();
			
			opXls.writeCell(13,2,viewVer);
			opXls.writeCell(14,2,severVer);
			opXls.writeCell(15,2,dbEdition);
			opXls.writeCell(16,2,dbVer);
			opXls.writeCell(20,2,regist);
		}
		
		/// <summary>
		/// Return the testcase name from runlist according to the order in rxtst
		/// </summary>
		/// <returns>String</returns>
		public string getTestCaseName()
		{
			if(m_Runlist.Count != 0) 
			{
				object[] arg = (object[])m_Runlist[0];
				return (string)arg[0];
			}
			
			return "";
		}

		/// <summary>
		/// After running one case, remove it from runlist
		/// </summary>
		/// <param name="testcaseName">testcaseName</param>
		public void runOverOneCase(string testcaseName)
		{
			m_Runlist.RemoveAt(0);
		}
		
		//**************************************************************************
		/// <summary>
		/// Comparer of the runlist, make sure the order in runlist is the order in rxtst
		/// </summary>
		/// <para> Author: Peter Yao</para>
		public class myRunlistCompareClass : IComparer
		{
			//**********************************************************************
			/// <summary>
			/// The order is accordint to the int value y, from small to large
			/// </summary>
     		int IComparer.Compare( Object x, Object y )  
      		{	
     			object[] argX = (object[])x;
     			object[] argY = (object[])y;
     			int iX = (int)argX[1];
     			int iY = (int)argY[1];
     			return( iX.CompareTo(iY) );
      		}
   		}

		/// <summary>
		/// Recursion to browse all rxtst test cases tree to generate running list.
		/// </summary>
		/// <param name="xnf">xnf</param>
		/// <param name="htAllCaseName">htAllCaseName</param>
		/// <param name="iNum">iNum</param>
		public void generateRunningList(XmlNode xnf, Hashtable htAllCaseName, ref int iNum)
		{
			if(xnf.HasChildNodes)
        	{
				//XmlNodeList childnodes = xnf.ChildNodes;    
				foreach(XmlNode childnode in xnf.ChildNodes)
				{
					generateRunningList(childnode,htAllCaseName,ref iNum);
				}
			}
			else
       		{        	   
				if(xnf.Name.Equals("testmodule"))
				{					
					XmlElement xeCurrentNode = (XmlElement)xnf;
					if(!xeCurrentNode.GetAttribute("name").Equals("TestCaseDriver"))
						return;
					XmlElement xe=(XmlElement)xnf.ParentNode;
       				htAllCaseName.Add(xe.GetAttribute("id"),
       		                  new object[] {xe.GetAttribute("name"), iNum} );
       				iNum ++;		
				}	
       		}        	
        	        			
		}
		
		/// <summary>
		/// Load the rxtst file, and get the runlist of testcase name
		/// </summary>
		/// <returns>runlist</returns>
		public ArrayList getRunlist()
		{
			ArrayList runlist = new ArrayList();
			
			string xmlpath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(),
                                                 "NformTester.rxtst");
        	XmlDocument xmldoc = new XmlDocument();
        	xmldoc.Load(xmlpath);
        	XmlNodeList nodes ;
        	XmlElement xe;
        	
        	Hashtable htAllCaseName = new Hashtable();
        	nodes = xmldoc.SelectNodes("/testsuite/content/testcase");
        	int iNum = 0;
        	
        	int abcddd = nodes.Count;
        	
        	foreach(XmlNode xnf in nodes)
        	{	
        		generateRunningList(xnf,htAllCaseName, ref iNum);        		        			
        	}
        	
        	
        	int abc = htAllCaseName.Count;
        	
        	nodes = xmldoc.SelectNodes("/testsuite/testconfigurations/testconfiguration/testcase");
        	for(int i=0; i< nodes.Count; i++)
        	{
        		xe=(XmlElement)nodes[i];
        		object[] arg = (object[])htAllCaseName[xe.GetAttribute("id")];
        		if(arg == null)
        			continue;
        		runlist.Add(arg);
        		
        	}
        	        	
        	IComparer myComparer = new myRunlistCompareClass();
        	runlist.Sort(myComparer);
        	
//        	MessageBox.Show(runlist.Count.ToString());
//        	
//        	for(int testNum = 0; testNum < runlist.Count; testNum++)
//        	{
//        		object[] arg = (object[])runlist[testNum];
//        		MessageBox.Show((string)arg[0] +"      "+arg[1].ToString());
//        	}
        	
        	return runlist;

		}
		
		/// <summary>
		/// Load the languagefile
		/// </summary>
		/// <returns>htName</returns>
		public Hashtable LoadLanguageInfo()
		{
			string currentLanguage = System.Globalization.CultureInfo.InstalledUICulture.Name.ToString();

            Hashtable m_LanguageFileMap = new Hashtable();
            m_LanguageFileMap.Add("zh-CN","CN");
            m_LanguageFileMap.Add("en","EN");

            string xmlpath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(),
            	                      "AppResource_"+m_LanguageFileMap[currentLanguage]+".xml");
            
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(xmlpath);
            XmlNodeList nodes;
            XmlElement xe;

            Hashtable htName = new Hashtable();
            nodes = xmldoc.SelectNodes("/Resource/Form/Controls/Control");
            int iNum = 0;
            foreach (XmlNode xnf in nodes)
            {
                xe = (XmlElement)xnf;
                htName.Add(xe.GetAttribute("name"),xe.GetAttribute("text"));
                iNum++;
            }
            
            return htName;
		}

		/// <summary>
		/// Return the string text according to LanguageID
		/// </summary>
		/// <param name="LanguageID">LanguageID</param>
		/// <returns>String</returns>
        public string GetLanguageString(string LanguageID)
        {
        	// return htLanguageInfo[LanguageID];
        	return "";
	    }
        
        /// <summary>
        ///  Create one file
        /// </summary>
        /// <param name="pathname">pathname</param>
		public void CreateFile(string pathname)
		{
			if(!Directory.Exists(pathname))
			{ 
				Directory.CreateDirectory(pathname);
			}
		}	
		
		//**********************************************************************
		/// <summary>
		/// Before excute, verify pre-condition.
		/// </summary>
		public void VerifyCondition()
		{
			//b_License=true, license is necessary;b_License=false, license is unnecessary;
			//b_Register=true, Register is necessary;b_Register=false, Register is unnecessary;
			//b_Device=true, many more devices are necessary;b_Device=false, many more devices are unnecessary;
			//b_Alarm=true, Alarm is necessary;b_Alarm=false, Alarm is unnecessary;
			bool b_License =getLicenseCP();
			bool b_Register = getRegisterCP();
			bool b_Device = getDeviceCP();
			bool b_Alarm = getAlarmCP();
			
			//Login
			NformRepository repo = NformRepository.Instance;
			repo.NFormApp.LogintoLiebertNformWindow.FormLogin_to_LiebertR_Nform.Username.PressKeys(m_strUserName);
			repo.NFormApp.LogintoLiebertNformWindow.FormLogin_to_LiebertR_Nform.Password.PressKeys(m_stPassword);
			repo.NFormApp.LogintoLiebertNformWindow.FormLogin_to_LiebertR_Nform.ServerCombo.PressKeys(m_strServerName);
			repo.NFormApp.LogintoLiebertNformWindow.FormLogin_to_LiebertR_Nform.Login.Click();
			Delay.Milliseconds(3000);
			
			if(repo.NFormApp.LicensesWindow.FormReminder.NoInfo.Exists())
			{
				repo.NFormApp.LicensesWindow.FormReminder.No.Click("53;10");
			}
			
			if(repo.NFormApp.LogintoLiebertNformWindow.FormEvaluation_Period_Expiration.OKInfo.Exists())
			{
				repo.NFormApp.LogintoLiebertNformWindow.FormEvaluation_Period_Expiration.OK.Click();
			}
			
			
			//Verify the license status
			repo.NFormApp.NformG2Window.FormMain.Configure.Click();
			repo.NFormApp.NformG2Window.FormMain.Licenses.Click();
			int license_count = repo.NFormApp.LicensesWindow.FormLicenses.LicenseList.Items.Count;
			if(license_count != 0)
			   Validate.AreEqual(true, b_License);
			else
				Validate.AreEqual(false, b_License);
			
			repo.NFormApp.LicensesWindow.FormLicenses.Close.Click();
			Delay.Milliseconds(1000);
			
			//Verify the Register status
			repo.NFormApp.NformG2Window.FormMain.Help.Click();
			repo.NFormApp.NformG2Window.FormMain.About_Liebert_Nform.Click();
			repo.NFormApp.Help.FormAbout_LiebertR_Nform.TabRegistration.Click();
			bool unregister_flag = repo.NFormApp.Help.FormAbout_LiebertR_Nform.RegistrationDscr.TextValue.Equals("Software assurance: Unregistered");
			if (unregister_flag == true)
				Validate.AreEqual(false, b_Register);
			else
				Validate.AreEqual(true, b_Register);
			repo.NFormApp.Help.FormAbout_LiebertR_Nform.OK.Click();
			Delay.Milliseconds(1000);

			//Verify the Device status
			repo.NFormApp.NformG2Window.FormMain.Configure.Click();
			repo.NFormApp.NformG2Window.FormMain.Devices.Click();
			int device_count = repo.NFormApp.ManagedDevicesWindow.FormManaged_Devices.Managed_device_table.Rows.Count;
			if(device_count > 1)
			   Validate.AreEqual(true, b_Device);
			else
				Validate.AreEqual(false, b_Device);
			repo.NFormApp.ManagedDevicesWindow.FormManaged_Devices.Close.Click();
			Delay.Milliseconds(1000);
	
			//Verify the Alarm status
			repo.NFormApp.NformG2Window.FormMain.Navigate.Click();
			repo.NFormApp.NformG2Window.FormMain.Alarms.Click();
			int alarm_count = repo.NFormApp.NformG2Window.FormMain.All_alarms_table.Rows.Count;
			if(alarm_count != 0)
			   Validate.AreEqual(true, b_Alarm);
			else
				Validate.AreEqual(false, b_Alarm);
			Delay.Milliseconds(1000);
		
			repo.NFormApp.NformG2Window.FormMain.Application.Click();
			repo.NFormApp.NformG2Window.FormMain.Login_As.Click();
			Delay.Milliseconds(1000);
		
		}	
		
		/// <summary>
		/// Get checkpoint of License
		/// </summary>
		/// <returns>Bool</returns>
		public bool getLicenseCP()
		{
			System.Xml.XmlDocument xdoc = new XmlDocument();
			xdoc.Load(m_strPrecondition);
			string License = xdoc.SelectSingleNode("configuration/License/text()").Value;
			bool b_License = false;
			return b_License = Convert.ToBoolean(License);
		}

		/// <summary>
		/// Get checkpoint of Register 
		/// </summary>
		/// <returns>Bool</returns>
		public bool getRegisterCP()
		{
			System.Xml.XmlDocument xdoc = new XmlDocument();
			xdoc.Load(m_strPrecondition);
			string Register = xdoc.SelectSingleNode("configuration/Register/text()").Value;
			bool b_Register = false;
			return b_Register = Convert.ToBoolean(Register);
		}
		
		/// <summary>
		/// Get checkpoint of Device
		/// </summary>
		/// <returns>Bool</returns>
		public bool getDeviceCP()
		{
			System.Xml.XmlDocument xdoc = new XmlDocument();
			xdoc.Load(m_strPrecondition);
			string Device = xdoc.SelectSingleNode("configuration/Device/text()").Value;
			bool b_Device = false;
			return b_Device = Convert.ToBoolean(Device);
		}
		
		/// <summary>
		/// Get checkpoint of Alarm
		/// </summary>
		/// <returns>Bool</returns>
		public bool getAlarmCP()
		{
			System.Xml.XmlDocument xdoc = new XmlDocument();
			xdoc.Load(m_strPrecondition);
			string Alarm = xdoc.SelectSingleNode("configuration/Alarm/text()").Value;
			bool b_Alarm = false;
			return b_Alarm = Convert.ToBoolean(Alarm);
		}
		
	}

}
