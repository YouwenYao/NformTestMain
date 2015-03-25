/*
 * Created by Ranorex
 * User: x93292
 * Date: 2013/5/8
 * Time: 14:40
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using Lextm.SharpSnmpLib;
using Lextm.SharpSnmpLib.Messaging;
using System.Net;
using System.Net.Sockets;
using Emerson.NetworkPower.Velocity;
using System.Collections;
using Microsoft.Win32;

namespace NformTester.lib
{
	/// <summary>
	/// Description of LxDeviceAvailable.
	/// </summary>
	public class LxDeviceAvailable
	{
		//**********************************************************************
		/// <summary>
		/// Constructer.
		/// </summary>
		public LxDeviceAvailable()
		{	
		}
		
		/// <summary>
		///Get available snmp device list.
		/// </summary>
		public void CheckSnmpDevice()
		{
			CheckSNMPDeviceAvailable();
		}
		
		/// <summary>
		///Get available velocity device list.
		/// </summary>
		public void CheckVelDevice()
		{
		//	setVelKeyName();
			CheckVelocityDeviceAvailable();
		}

		
		/// <summary>
		/// Check the SNMP devices in the Devices.ini are available or not.
		/// If some devices are unavailable, the Devices.ini should be changed to give
		/// tester all available devices.
		/// </summary>
		/// <returns>list</returns>
		public static List<String> CheckSNMPDeviceAvailable(){			
			
		    List<string> notAvailable = new List<string>();
			int timeout = 3;
			VersionCode version = VersionCode.V1;
			IPEndPoint endpoint = new IPEndPoint(IPAddress.Parse("1.1.1.1"),161);
			OctetString community = new OctetString("public");
            
            ObjectIdentifier objectId = new ObjectIdentifier("1.3.6.1.2.1.11.1.0");
            Variable var = new Variable(objectId);   
            IList<Variable> varlist = new System.Collections.Generic.List<Variable>();
            varlist.Add(var);
            Variable data;
            IList<Variable> resultdata;
            IDictionary<string, string> AllDeviceInfo = new Dictionary<string, string> ();
            IDictionary<string, string> SNMPDeviceInfo = new Dictionary<string, string> ();
            
            AllDeviceInfo = AppConfigOper.mainOp.DevConfigs;

			foreach(string key in AllDeviceInfo.Keys)
			{
				if(key.ToUpper().StartsWith("SNMP"))
				{
					SNMPDeviceInfo.Add(key, AllDeviceInfo[key]);
				}
			}
			
			foreach (KeyValuePair<string, string>device in SNMPDeviceInfo)
			{ 
				Console.WriteLine("SNMPDeviceInfo: key={0},value={1}", device.Key, device.Value);
			}
			
            foreach(string deviceIp in SNMPDeviceInfo.Values)
            {
            	try
            	{
            		endpoint.Address = IPAddress.Parse(deviceIp);
	            	resultdata = Messenger.Get(version,endpoint,community,varlist,timeout);
	            	data = resultdata[0];
	            	Console.WriteLine("The device:" + deviceIp + "is availabe");
            	}
            	catch(Exception ex)
	            {
	            	notAvailable.Add(deviceIp);
	            	Console.WriteLine("There is no device in this ip address."+ deviceIp);
	            	string log = ex.ToString();
	            	continue;
	            }	
            }
            return notAvailable;
	   }

		/// <summary>
		/// Check the velocity devices in the Devices.ini are available or not.
		/// If some devices are unavailable, the Devices.ini should be changed to give
		/// tester all available devices.
		/// </summary>
		/// <returns>list</returns>
		public static List<String> CheckVelocityDeviceAvailable(){
			// There are all devices in Device.ini.
			List<string> notAvailable = new List<string>();		
            IDictionary<string, string> AllDeviceInfo = new Dictionary<string, string> ();
            IDictionary<string, string> VelDeviceInfo = new Dictionary<string, string> ();
			AllDeviceInfo = AppConfigOper.mainOp.DevConfigs;

			foreach(string key in AllDeviceInfo.Keys)
			{
				if(key.ToUpper().StartsWith("VELOCITY"))
				{
					VelDeviceInfo.Add(key, AllDeviceInfo[key]);
				}
			}
			
			foreach (KeyValuePair<string, string>device in VelDeviceInfo)
			{ 
				Console.WriteLine("VelDeviceInfo: key={0},value={1}", device.Key, device.Value);
			}
            
            
            
	        string NformPath = "";
			RegistryKey Key;
			Key = Registry.LocalMachine;
			//Nform in Register table
			// The path is based on your pc
		//  RegistryKey myreg = Key.OpenSubKey("software\\Liebert\\Nform");
			RegistryKey myreg = Key.OpenSubKey("software\\Wow6432Node\\Liebert\\Nform");
			NformPath = myreg.GetValue("InstallDir").ToString();
			myreg.Close();
			Console.WriteLine("NformPath:" + NformPath);
			
			 ClientEngine ve = ClientEngine.GetClientEngine();
	         string GDDpath = System.IO.Path.Combine(NformPath,@"bin\mds");
	         System.IO.Directory.CreateDirectory(GDDpath);
	         int startupflag = ve.Startup(GDDpath);
	         Console.WriteLine("GDDPath:" + GDDpath);
	         Console.WriteLine("startupflag:" + startupflag);
	         
	         CommsChannel channel = null;
	     
	         try
	         {
	         	channel = ve.OpenChannel(CommStack.BACnetIp);
	         }
	         catch(Exception e)
	         {
	         	Console.WriteLine("Can not open this channel! because: " + e.StackTrace.ToString());
	         }
	         
		     foreach(string deviceIp in VelDeviceInfo.Values)
            {
            	if(channel != null)
            	{
            		DeviceNode node = channel.StrToEntity(deviceIp);
            		if(node != null)
			     	{
			     		Console.WriteLine("The device:" + deviceIp + " is availabe");
			     	}
			     	else
			     	{
			     		notAvailable.Add(deviceIp);
			     		Console.WriteLine("There is no device in this ip address."+ deviceIp + "!" );
			     	}
            	}
            	else
			    {
			     	Console.WriteLine("This channel is not available!");
			    }
		     }
		     channel.Dispose();
		     return notAvailable;
		}
   }
	
}