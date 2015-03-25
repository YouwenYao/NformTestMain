/*
 * Created by Ranorex
 * User: y93248
 * Date: 2011-11-21
 * Time: 14:41
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Reflection;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Collections;
using System.Xml;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;
using Ranorex.Core.Repository;

namespace NformTester.lib
{
	//**************************************************************************
	/// <summary>
	/// Define a data structure for a command line in scritps
	/// </summary>
	/// <para> Author: Peter Yao</para>
	public class LxScriptItem
	{
		/// <summary>
		/// Index of command
		/// </summary>
		public string m_Index;
		
		/// <summary>
		/// Type of command
		/// </summary>
		public string m_Type;
		
		/// <summary>
		/// WindowName, which form this command work on
		/// </summary>
		public string m_WindowName;
		
		/// <summary>
		/// Component, which component this command work on
		/// </summary>
		public string m_Component;
		
		/// <summary>
		/// Action of this command
		/// </summary>
		public string m_Action;
		
		/// <summary>
		/// Arguments of this command
		/// </summary>
		public string m_Arg1;
		/// <summary>
		/// Arguments of this command
		/// </summary>
		public string m_Arg2;
		/// <summary>
		/// Arguments of this command
		/// </summary>
		public string m_Arg3;
		/// <summary>
		/// Arguments of this command
		/// </summary>
		public string m_Arg4;
		/// <summary>
		/// Arguments of this command
		/// </summary>
		public string m_Arg5;
		/// <summary>
		/// Arguments of this command
		/// </summary>
		public string m_Arg6;
		
		/// <summary>
		/// mainOp
		/// </summary>
		public static LxSetup mainOp = LxSetup.getInstance();
		
		/// <summary>
		/// Devices in DeviceConfig.xml
		/// </summary>
		public static Hashtable DeviceInfo = new Hashtable();
		
		
		/// <summary>
		/// Get the repository instance
		/// </summary>
		public static NformRepository repo = NformRepository.Instance;
		
		//**********************************************************************
		/// <summary>
		/// Constructer.
		/// </summary>
		public LxScriptItem()
		{
		}
		
		/// <summary>
		/// If this command has arguments, then return true.
		/// If this command has no arguments, then return false.
		/// </summary>
		/// <returns>true/false</returns>
		public bool hasArg()
		{
			if (m_Arg1 == "" && m_Arg2 == "" && m_Arg3 == ""
			   	&& m_Arg4 == "" && m_Arg5 == "" && m_Arg6 == "")
			{
				return false;			
			}
			return true;
		}
		
		/// <summary>
		/// git value from app.config for script. such as this format: $Ip_address$
		/// </summary>
		/// <param name="name">name</param>
		/// <returns>value</returns>
		public string parseToValue(string name)
        {
			return mainOp.parseToValue(name);
        }
		

		/// <summary>
		/// If argument has text, then remove symbol"
		/// </summary>
		/// <returns>String</returns>
		public string getArgText()
		{
			return parseToValue(m_Arg1);
		}

		/// <summary>
		/// Get argument2
		/// </summary>
		/// <returns>String</returns>
		public string getArg2Text()
		{
			return parseToValue(m_Arg2);
		}

		/// <summary>
		/// Get argument3
		/// </summary>
		/// <returns>String</returns>
		public string getArg3Text()
		{
			return parseToValue(m_Arg3);
		}
		
		/// <summary>
		/// Get argument4
		/// </summary>
		/// <returns>String</returns>
		public string getArg4Text()
		{
			return parseToValue(m_Arg4);
		}

		/// <summary>
		/// Get argument5
		/// </summary>
		/// <returns>String</returns>
		public string getArg5Text()
		{
			return parseToValue(m_Arg5);
		}

		/// <summary>
		/// Get argument6
		/// </summary>
		/// <returns>String</returns>
		public string getArg6Text()
		{
			return parseToValue(m_Arg6);
		}	
		
		
		/// <summary>
		/// According to the name of windows and component, find the componentinfo
		/// object in repository
		/// </summary>
		/// <returns>RepoItemInfo</returns>
		public RepoItemInfo getComponentInfo()
		
		{
			string windowsName = m_WindowName;
			string componentName = m_Component;
			Type objType = repo.GetType();
           
           	object obj = repo;
           	PropertyInfo pi = objType.GetProperty("NFormApp");
           	obj = pi.GetValue(repo,null);
           	objType = obj.GetType();
           	PropertyInfo[] piArrLev1 = objType.GetProperties(BindingFlags.Instance | BindingFlags.Public);
       	   	foreach (PropertyInfo piLev1 in piArrLev1)
        	{	
       	   		if(piLev1.Name.CompareTo("UseCache") == 0)
       	   		{
       	   			continue;
       	   		}
       	   		object objLogicGroup = piLev1.GetValue(obj,null);
       	   		objType = objLogicGroup.GetType();
       	   		PropertyInfo[] piArrLev2 = objType.GetProperties(BindingFlags.Instance | BindingFlags.Public);
       	   		foreach (PropertyInfo piLev2 in piArrLev2)
        		{
       	   			if(piLev2.Name.CompareTo(windowsName) == 0)
       	   			{
       	   				object objWindows = piLev2.GetValue(objLogicGroup,null);
       	   				objType = objWindows.GetType();
       	   				PropertyInfo[] piArrComp = objType.GetProperties(BindingFlags.Instance | BindingFlags.Public);
       	   				foreach (PropertyInfo piCom in piArrComp)
        				{
       	   					if(piCom.Name.CompareTo(componentName + "Info") == 0)
       	   					{
       	   						RepoItemInfo objComponets = (RepoItemInfo)piCom.GetValue(objWindows,null);
       	   						return objComponets;
       	   					} 
       	   				} 
       	   			} 
       	   		}  
        	} 
       	   	
       	   	return null;
		}
		
		/// <summary>
		/// According to the name of windows and component, find the component
		/// object in repository
		/// </summary>
		/// <returns>Object</returns>
		public Object getComponent()
		{
			string windowsName = m_WindowName;
			string componentName = m_Component;
			Type objType = repo.GetType();

           	object obj = repo;
           	PropertyInfo pi = objType.GetProperty("NFormApp");
           	obj = pi.GetValue(repo,null);
           	objType = obj.GetType();
           	PropertyInfo[] piArrLev1 = objType.GetProperties(BindingFlags.Instance | BindingFlags.Public);
       	   	foreach (PropertyInfo piLev1 in piArrLev1)
        	{	
       	   		if(piLev1.Name.CompareTo("UseCache") == 0)
       	   		{
       	   			continue;
       	   		}
       	   		object objLogicGroup = piLev1.GetValue(obj,null);
       	   		objType = objLogicGroup.GetType();
       	   		PropertyInfo[] piArrLev2 = objType.GetProperties(BindingFlags.Instance | BindingFlags.Public);
       	   		foreach (PropertyInfo piLev2 in piArrLev2)
        		{
       	   			if(piLev2.Name.CompareTo(windowsName) == 0)
       	   			{
       	   				object objWindows = piLev2.GetValue(objLogicGroup,null);
       	   				objType = objWindows.GetType();
       	   				PropertyInfo[] piArrComp = objType.GetProperties(BindingFlags.Instance | BindingFlags.Public);
       	   				foreach (PropertyInfo piCom in piArrComp)
        				{
       	   					if(piCom.Name.CompareTo(componentName) == 0)
       	   					{
       	   						object objComponets = piCom.GetValue(objWindows,null);
       	   						return objComponets;
       	   					} 
       	   				} 
       	   			} 
       	   		}  
         		
        	} 
       	   	
       	   	return null;
		}
		
	}
}
