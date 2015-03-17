/*
 * Created by Ranorex
 * User: x93292
 * Date: 2015-01-15
 * Time: 10:21
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Data.SqlClient;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Collections.Generic;

namespace NformTester.lib
{
	/// <summary>
	/// Description of LxSQLDbOperation.
	/// </summary>
	public class LxSQLDbOperation:LxDBOper
	{
	
		/// <summary>
		/// Get table value from database
		/// </summary>
		public List<string[]> GetTableValue(SqlConnection conn, string queryStr)
		{
			if (conn.State == ConnectionState.Closed)
				conn.Open();
			List<string[]> queryResult = new List<string[]>();
			DataSet ds = new DataSet();
			try
			{
				SqlDataAdapter adapter = new SqlDataAdapter(queryStr, conn);
				adapter.Fill(ds,"result");
				int colcount = ds.Tables["result"].Columns.Count;
				Console.WriteLine("the colcount is "+colcount);
				//Parse DataSet to List, be used to collect result data.
				queryResult = DataSetToList(ds);
			}
			catch(Exception ex)
			{
				Console.WriteLine("the Error in GetVarLenColumnSize is "+ex.StackTrace.ToString());
			}
			
			return queryResult;
			
		}
		
		/// <summary>
		/// Get size of database
		/// </summary>
		/// <param name="conn">conn</param>
		/// <param name="dbName">Nform, NformAlm, NformLog</param>
		/// <returns>DbSize</returns>
		public string GetDbSize(SqlConnection conn, string dbName)
		{
			string Dbsize = "";
			string sizeDB = @"sp_helpdb;";
			DataSet ds = new DataSet();
			if (conn.State == ConnectionState.Closed)
				conn.Open();
			
	        SqlDataAdapter adapt = new SqlDataAdapter(sizeDB, conn);
	        adapt.Fill(ds,"sizetable");
	        
	        int recordcount = ds.Tables["sizetable"].Rows.Count;
	        for(int i=0;i<recordcount;i++)
	        {
	        	if((ds.Tables["sizetable"].Rows[i][0]).ToString().Equals(dbName))
	        	   {
	        			Dbsize = (ds.Tables["sizetable"].Rows[i][1]).ToString();
		        	   	break;
	        	   }
	        }
	        conn.Close();
	        Console.WriteLine( dbName + "The DB size is " + Dbsize);
			return Dbsize;
		}
		
		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		public bool IncreaseAlarmDbSize()
		{
			GenerateData("AlarmClearLogData");
			GenerateData("AlarmGenerateLogData");
			return true;
		}
		
		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		public bool IncreaseDatalogDbSize()
		{
			GenerateData("DataLogClearLogData");
			GenerateData("DataLogGenerateLogData");
			return true;
		}
		
		/// <summary>
		/// 
		/// </summary>
		/// <param name="commandstr"></param>
		public static void GenerateData(string commandstr)
		{
			string strApplicationName = AppConfigOper.getConfigValue(commandstr);
			Console.WriteLine("Command is: " + strApplicationName);
			Process clrpro = new Process();
			FileInfo clrfile = new FileInfo(strApplicationName);
			clrpro.StartInfo.WorkingDirectory = clrfile.Directory.FullName;
			clrpro.StartInfo.FileName = strApplicationName;
			clrpro.StartInfo.CreateNoWindow = false;
			clrpro.Start();
			clrpro.WaitForExit();
		}
		
	}
}
