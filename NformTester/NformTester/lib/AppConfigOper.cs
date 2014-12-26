/*
 * Created by Ranorex
 * User: x93292
 * Date: 2014-12-19
 * Time: 14:40
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.IO;
using System.Configuration;
using System.Collections;
using System.Xml;

namespace NformTester.lib
{
	/// <summary>
	/// This class is used to operate on app.config.
	/// </summary>
	public class AppConfigOper
	{
		/// <summary>
		/// LxSetup class
		/// </summary>
		public static LxSetup mainOp = LxSetup.getInstance();
		
		/// <summary>
		/// Devices in DeviceConfig.xml
		/// </summary>
		public static Hashtable DeviceInfo = new Hashtable();
		
		/// <summary>
		/// Constructer
		/// </summary>
		public AppConfigOper()
		{
		}
		
		/// <summary>
		/// git value from app.config for script. such as this format: $Ip_address$
		/// </summary>
		/// <param name="name">name</param>
		/// <returns>value</returns>
		public static string parseToValue(string name)
        {
            if (name.Equals(""))
            {
                return "";
            }
			var configs = mainOp.configs;
            string addr = name;
            if (name.Substring(0, 1) == "$" && name.Substring(name.Length - 1, 1) == "$")
            {
                string key = name.Substring(1, name.Length - 2);
                LxIniFile confFile = new LxIniFile(System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(),
                                                 "Devices.ini"));
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
                
                confFile = new LxIniFile(System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(),
                                                 "UsedDevices.ini"));
                confFile.WriteString("AvailableDevices",key,result);
            }

            return addr.Replace("\"","");
        }
		
		/// <summary>
		/// Get value of key from app.config
		/// </summary>
		/// <param name="key">key</param>
		/// <returns>value</returns>
		public static string getConfigValue(string key)
		{
          	var configs = mainOp.configs;
          	string value = null;
            if(configs.ContainsKey(key))
            {
               value = configs[key];
            }
            else
            {
              	value = configs["Default"];
            }
	      	return value;
		}

		/// <summary>
		/// get device info, and set value to DeviceInfo.
		/// </summary>
		/// <param name="DeviceType">DeviceType</param>
		/// <param name="ConfigPath">ConfigPath</param>
		/// <returns>Hashtable</returns>
		public static Hashtable GetDevicesInfo(string DeviceType, string ConfigPath)
		{
			ArrayList runlist = new ArrayList();
			Hashtable DeviceInfo = new Hashtable();
			string xmlpath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(),
                                                 ConfigPath);
			XmlDocument xmldoc = new XmlDocument();
        	xmldoc.Load(xmlpath);
        	XmlNodeList nodes;
        	
        	if(DeviceType.Equals("snmp"))
        		nodes = xmldoc.SelectNodes("/configuration/snmp");
        	else
        		nodes = xmldoc.SelectNodes("/configuration/velocity");
        	foreach(XmlNode xnf in nodes)
        	{	
        		generateDeviceList(xnf,DeviceInfo);        		        			
        	}
        	return DeviceInfo;
		}
		

		/// <summary>
		/// get device ip
		/// </summary>
		/// <param name="xnf">xnf</param>
		/// <param name="DeviceInfo">DeviceInfo</param>
		public static void generateDeviceList(XmlNode xnf,Hashtable DeviceInfo)
		{
			if(xnf.HasChildNodes)
        	{
				//XmlNodeList childnodes = xnf.ChildNodes;    
				foreach(XmlNode childnode in xnf.ChildNodes)
				{
					generateDeviceList(childnode,DeviceInfo);
				}
			}
			else
       		{        	   
				if(xnf.Name.Equals("device"))
				{					
					XmlElement xe = (XmlElement)xnf;
       				DeviceInfo.Add(xe.GetAttribute("name"),xe.GetAttribute("ip"));	
				}	
       		}        	
        	        			
		}
		
	}
}
