/*
 * Created by Ranorex
 * User: x93292
 * Date: 2015-01-05
 * Time: 11:42
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Data.SqlServerCe;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;

namespace NformTester.lib
{
	/// <summary>
	/// Description of LxCEDbOperation.
	/// </summary>
	public class LxCEDbOperation:LxDBOper
	{
		public LxCEDbOperation()
		{
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
		/// <param name="dbName"></param>
		/// <param name="queryStr"></param>
		/// <returns></returns>
		public List<string[]> GetTableValue(string dbName, string queryStr)
		{
			string connStr = GetDBConnString(dbName);
			List<string[]> queryResult = new List<string[]>();
			using(SqlCeConnection conn = new SqlCeConnection(connStr))
			{
				DataSet ds = new DataSet();
				try
				{
					SqlCeDataAdapter adapter = new SqlCeDataAdapter(queryStr, conn);
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
			}
			return queryResult;
			
		}
		
		/// <summary>
		/// Set table’s value (To verify the value from Nform interface)
		/// </summary>
		/// <param name="dbname"></param>
		/// <param name="tablename"></param>
		/// <param name="value"></param>
		/// <returns></returns>
		public bool SetTableValue(string dbname,string tablename,string value)
		{
			return true;
		}
		/// <summary>
		/// Get size of database
		/// </summary>
		/// <returns></returns>
		public double GetAlarmDbSize()
		{
			 //define some const
			const int  SQLCE_FIXED_BYTES_PER_ALARM_REC = 272;
			double TotalCnt = 0;
			double RecCnt = 0;
			double VarLenColumnSize = 0;
			double VeriableLength = 0;
			double DataSize = 0;
			string connStr = GetDBConnString("NformAlm.sdf");
			using(SqlCeConnection conn = new SqlCeConnection(connStr))
			{
				//Open the connection of sqlce database
				if (conn.State == ConnectionState.Closed)
				conn.Open();
		//		Console.WriteLine("The state of conn is:"+conn.State.ToString());
				//step1 TotalCnt
				TotalCnt = GetTotalCnt(conn,"Alarm");   
				//step2 RecCnt
				RecCnt = GetRecCnt(conn,"Alarm");
		        //step3 VarLenColumnSize
		        VarLenColumnSize = GetVarLenColumnSize(conn,"Alarm");
			}
			
	//		Console.WriteLine("The TotalCnt in GetAlarmDbSize is:"+TotalCnt);
	//		Console.WriteLine("The RecCnt in GetAlarmDbSize is:"+RecCnt);
	//		Console.WriteLine("The VarLenColumnSize in GetAlarmDbSize is:"+VarLenColumnSize);
	        //step4 VeriableLength = VarLenColumnSize / RecCnt
	        VeriableLength = VarLenColumnSize / RecCnt;
	//      Console.WriteLine("The VeriableLength in GetAlarmDbSize is:"+VeriableLength);
	        //step5 DataSize = TotalCnt * (SQLCE_FIXED_BYTES_PER_ALARM_REC + (VeriableLength * 2) + 6) / 1048576D <MB>
	        DataSize = TotalCnt * (SQLCE_FIXED_BYTES_PER_ALARM_REC + (VeriableLength * 2) + 6) / 1048576D;
	        Console.WriteLine("The DataSize in GetAlarmDbSize is:"+DataSize);
			return DataSize;
		}
		
		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		public double GetDataLogDbSize()
		{	
			//connect sqlce NformLog.sdf
	        string connStr = GetDBConnString("NformLog.sdf");	
	        //define some const
			const int  SQLCE_FIXED_BYTES_PER_DATA_LOG_REC = 70;
			double TotalCnt = 0;
			double RecCnt = 0;
			double VarLenColumnSize = 0;
			double VeriableLength = 0;
			double DataSize = 0;
			using(SqlCeConnection conn = new SqlCeConnection(connStr))
			{
				//Open the connection of sqlce database
				if (conn.State == ConnectionState.Closed)
				conn.Open();
		//		Console.WriteLine("The state of conn is:"+conn.State.ToString());
				//step1 TotalCnt
				TotalCnt = GetTotalCnt(conn,"DataLog");
				//step2 RecCnt
				RecCnt = GetRecCnt(conn,"DataLog");
		        //step3 VarLenColumnSize
		        VarLenColumnSize = GetVarLenColumnSize(conn,"DataLog");
			}
			
	//		Console.WriteLine("The TotalCnt in GetDataLogDbSize is:"+TotalCnt);
	//		Console.WriteLine("The RecCnt in GetDataLogDbSize is:"+RecCnt);
	//		Console.WriteLine("The VarLenColumnSize in GetDataLogDbSize is:"+VarLenColumnSize);
	        //step4 VeriableLength = VarLenColumnSize / RecCnt
	        VeriableLength = VarLenColumnSize / RecCnt;  
	//      Console.WriteLine("The VeriableLength in GetDataLogDbSize is:"+VeriableLength);
	        //step5 DataSize = TotalCnt * (SQLCE_FIXED_BYTES_PER_ALARM_REC + (VeriableLength * 2) + 6) / 1048576D <MB>
	        DataSize = TotalCnt * (SQLCE_FIXED_BYTES_PER_DATA_LOG_REC + (VeriableLength * 2) + 6) / 1048576D;
	        Console.WriteLine("The DataSize in GetDataLogDbSize is:"+DataSize);
			return DataSize;
		}
		
		/// <summary>
		/// 
		/// </summary>
		/// <param name="dbName"></param>
		/// <returns></returns>
		public string GetDBConnString(string dbName)
		{
			string ConnectionString = @"Data Source=C:\Nform\db\"+dbName;
			return ConnectionString;
			
		}
		
		/// <summary>
		/// 
		/// </summary>
		/// <param name="tableName"></param>
		/// <returns></returns>
		public double GetTotalCnt(SqlCeConnection conn, string tableName)
		{
			//TotalCnt
			string cmdTotal = @"SELECT COUNT(*) AS TotalCnt FROM "+tableName+";";
		//	Console.WriteLine("The cmdTotal in GetTotalCnt is:"+cmdTotal);
			double TotalCnt = 0;
			
			try
			{
				SqlCeCommand cmd = new SqlCeCommand(cmdTotal, conn);
				Object queryAll = cmd.ExecuteScalar();
				TotalCnt = double.Parse(queryAll.ToString());
		 //		Console.WriteLine("the queryAll is "+queryAll.ToString());
			}
			catch(Exception ex)
			{
				Console.WriteLine("the Error in GetTotalCnt is "+ex.StackTrace.ToString());
				
			}
	//		Console.WriteLine("the TotalCnt is "+TotalCnt);
	        return TotalCnt;
		}
		
		/// <summary>
		/// 
		/// </summary>
		/// <param name="tableName"></param>
		/// <returns></returns>
		public double GetVarLenColumnSize(SqlCeConnection conn,string tableName)
		{
			//step3 VarLenColumnSize
			double VarLenColumnSize = 0; 
			DataSet ds = new DataSet();
			string cmdAll = @"SELECT * FROM "+tableName;
		//	Console.WriteLine("The cmdAll in GetVarLenColumnSize is:"+cmdAll);
			try
			{
				SqlCeDataAdapter adapter = new SqlCeDataAdapter(cmdAll, conn);
				adapter.Fill(ds,"alarm");
				int colcount = ds.Tables["alarm"].Columns.Count;
			//	Console.WriteLine("The colcount in GetVarLenColumnSize is:"+colcount);
		        for(int i=0;i<colcount;i++)
		        {
		        	string tmpname = ds.Tables["alarm"].Columns[i].ColumnName;	        
		        	string tmpsql = @"SELECT SUM(LEN(" + tmpname + ")) AS lens FROM (Select top(1000) * from "+tableName+") AS A;";
		        //	Console.WriteLine("the tmpsql in GetVarLenColumnSize is:" + tmpsql);
		        	SqlCeCommand tmpcmd = new SqlCeCommand(tmpsql, conn);
		        	Object queryAll = tmpcmd.ExecuteScalar();
		        	double lens = double.Parse(queryAll.ToString());
		        //	Console.WriteLine("the lens in GetVarLenColumnSize is:" + lens);
		        	VarLenColumnSize = VarLenColumnSize + lens;
				}
			}
			catch(Exception ex)
			{
				Console.WriteLine("the Error in GetVarLenColumnSize is "+ex.StackTrace.ToString());
			}
	//		Console.WriteLine("the VarLenColumnSize in GetVarLenColumnSize is:" + VarLenColumnSize);
	        return VarLenColumnSize;
		}
		
		/// <summary>
		/// 
		/// </summary>
		/// <param name="tableName"></param>
		/// <returns></returns>
		public double GetRecCnt(SqlCeConnection conn,string tableName)
		{
			double RecCnt = 0;
			string cmdstr = @"SELECT COUNT(*) AS RecCnt From (SELECT TOP (1000) * FROM "+tableName+") AS A;";
		//	Console.WriteLine("The cmdstr in GetRecCnt is:"+cmdstr);
			try
			{
				SqlCeCommand cmd = new SqlCeCommand(cmdstr, conn);
				Object queryAll = cmd.ExecuteScalar();
				RecCnt = double.Parse(queryAll.ToString());
		// 		Console.WriteLine("the queryAll is "+queryAll.ToString());
			}
			catch(Exception ex)
			{
				Console.WriteLine("the Error in GetRecCnt is "+ex.StackTrace.ToString());
				
			}
	//		Console.WriteLine("the GetRecCnt is "+RecCnt);
	        return RecCnt;
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