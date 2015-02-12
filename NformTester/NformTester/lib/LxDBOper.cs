/*
 * Created by Ranorex
 * User: x93292
 * Date: 2013/5/9
 * Time: 10:19
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Collections.Generic;

namespace NformTester.lib
{
	/// <summary>
	/// Be used to backup database before run scripts. And be used to restore database if scripts are failed.
	/// </summary>
	public class LxDBOper
	{
		/// <summary>
		///DB_DbType=1, Bundled database is used; 
		///DB_DbType=2, SQL Server database is used.
		/// </summary>
		public int DB_DbType;
		
		/// <summary>
		///BackUp Result. 
		/// </summary>
		public bool BackUpResult=false;
		
		/// <summary>
		///Restore Result. 
		/// </summary>
		public bool RestoreResult=false;
		
		/// <summary>
		/// Constructor
		/// </summary>
		public LxDBOper()
		{
		}
	
/*		
		/// <summary>
		/// Get database type from device.ini.
		/// Set DB_DbType.
		/// </summary>
		public void SetDbType()
		{
	        string groupName="Database";
	        string key="DB_DbType";
	        string DbType = ParseToValue(groupName,key);
	        DB_DbType = (int.Parse(DbType));
		}

	*/	
		
		/// <summary>
		/// Get database type from device.ini.
		/// Set DB_DbType.
		/// </summary>
		public void SetDbType(string DbType)
		{
	        DB_DbType = (int.Parse(DbType));
		}
		
		
		/// <summary>
		/// Get DB_DbType
		/// </summary>
		/// <returns>DB_DbType</returns>
		public int GetDbType()
		{
			return DB_DbType;
		}
		
		/// <summary>
		/// Get BackUpResult.
		/// </summary>
		/// <returns>BackUpResult</returns>
		public bool GetBackUpResult()
		{
			return BackUpResult;
		}
		
		/// <summary>
		/// Get RestoreResult.
		/// </summary>
		/// <returns>RestoreResult</returns>
		public bool GetRestoreResult()
		{
			return RestoreResult;
		}
		
		/*
		/// <summary>
		/// Parse the value from Devices.ini.
		/// </summary>
		/// <param name="GroupName">GroupName</param>
		/// <param name="key">key</param>
		/// <returns>result</returns>
		public string ParseToValue(string GroupName, string key)
        {
		  LxIniFile confFile = new LxIniFile(System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(),
                                                 "Devices.ini"));
          string def = confFile.GetString(GroupName, "Default", "null");
          string result = confFile.GetString(GroupName, key, def);
          return result;
        }
		
		*/
		/// <summary>
		/// Copy file.
		/// </summary>
		/// <param name="srcPath">srcPath</param>
		/// <param name="aimPath">aimPath</param>
	   public void CopyDir(string srcPath, string aimPath)
      {
        try
        {
            if (aimPath[aimPath.Length - 1] != System.IO.Path.DirectorySeparatorChar)
            {
                aimPath += System.IO.Path.DirectorySeparatorChar;
            }
            if (!System.IO.Directory.Exists(aimPath))
            {
                System.IO.Directory.CreateDirectory(aimPath);
            }
            string[] fileList = System.IO.Directory.GetFileSystemEntries(srcPath);
            foreach (string file in fileList)
            {
                if (System.IO.Directory.Exists(file))
                {
                    CopyDir(file, aimPath + System.IO.Path.GetFileName(file));
                } 
                else
                {
                    System.IO.File.Copy(file, aimPath + System.IO.Path.GetFileName(file), true);
                }
            }
        }
        catch (Exception e)
        {
        	throw e;
        }
    }
		
		/// <summary>
		/// Open the SQL server connection.
		/// </summary>
		/// <param name="conn">conn</param>
		public void OpenConnection(SqlConnection conn)
		{
		    try
          {
		    	conn.Open();
          }
	        catch(Exception ex)
	        {
	        	string log = ex.ToString();
	        	 Console.WriteLine(log);
	        }
       }
		
		
		/// <summary>
		/// Get size of database
		/// </summary>
		/// <returns></returns>
		public int GetDbSize(SqlConnection conn)
		{
			conn.Open();
			return 0;
		}
		
		/// <summary>
		/// Increase database size (we may integrate code or call the tool directly)
		/// </summary>
		/// <param name="conn"></param>
		/// <returns></returns>
		public bool IncreaseDbSize(SqlConnection conn)
		{
			conn.Open();
			return true;
		}
		
		
		/// <summary>
		/// Close the SQL server connection.
		/// </summary>
		/// <param name="conn">conn</param>
		public void CloseConnection(SqlConnection conn)
		{
		    try
          {
		    	conn.Close();
          }
		    catch(Exception ex)
	        {
	        	string log = ex.ToString();
	        }
       }
		
		/// <summary>
		/// Get the database connection strin from Devices.ini.
		/// </summary>
		/// <returns>connString</returns>
		public string GetDBConnString()
		{
			string DS = AppConfigOper.getConfigValue("DB_SQL_Server_Name");
			string UN = AppConfigOper.getConfigValue("DB_SQL_User_Name");
			string PWD = AppConfigOper.getConfigValue("DB_SQL_Password");
	        string connString = "Data Source="+DS+";"+"Initial Catalog=master;"+"User ID="+UN+";"+"Password="+PWD+";";
	        return connString;
		}
			
		
		/// <summary>
		/// Back up for bundled database;
		/// </summary>
		public void BackUpBundledDataBase()
		{
			  bool result = false;
			  Console.WriteLine("*****Start to back up bundled Database*****");
			  string sourceDBPath = AppConfigOper.getConfigValue("DB_Bundled_Path");
			  string targetDBPath = AppConfigOper.getConfigValue("DB_Bundled_Backup_Path");
	          
			   if (!Directory.Exists(sourceDBPath))
			   {
			   		BackUpResult = false;
			   		return;
			   }
		       //First delete the exsited directory.
			   if (Directory.Exists(targetDBPath))
			   {
			   	Directory.Delete(targetDBPath,true);
		       }
			   //Then create new directory with the same name.
		       Directory.CreateDirectory(targetDBPath);
		       //Restore the files.
		       
		       CopyDir(sourceDBPath,targetDBPath);
		       result = true;
		       Console.WriteLine("*****Finish to back up bundled Database*****");
		       BackUpResult = result;
		}
		
		/// <summary>
		/// Backpup database for SQL Server.
		/// </summary>
		/// <param name="conn">conn</param>
		public void ExcuteBakupQuery(SqlConnection conn)
        {	
			bool result = false;
			string backupPath = @"C:\Nform\SqlBackup";
			if (!System.IO.Directory.Exists(backupPath))
            {
				System.IO.Directory.CreateDirectory(backupPath);
            }
			if (conn.State == ConnectionState.Closed)
				conn.Open();
			string backupDB = @"Backup Database Nform To disk= 'C:\Nform\SqlBackup\Nform.bak';";
	        SqlCommand cmd = new SqlCommand(backupDB, conn);
	        cmd.ExecuteNonQuery();
	        backupDB =  @"Backup Database Nform To disk= 'C:\Nform\SqlBackup\NformAlm.bak';";
	        cmd.CommandText = backupDB;
	        cmd.ExecuteNonQuery();
	        backupDB =  @"Backup Database Nform To disk= 'C:\Nform\SqlBackup\NformLog.bak';";
	        cmd.CommandText = backupDB;
	        cmd.ExecuteNonQuery();     
	        result = true;
	        BackUpResult = result;
       }
		
	    /// <summary>
		/// Back up for SQL Server database;
		/// </summary>
		public void BackUpSQLServerDataBase()
		{
			Console.WriteLine("*****Start to back up SQL Server Database*****");
			SqlConnection conn = new SqlConnection();
			conn.ConnectionString = @"Data Source=NFORMTES-6FD309\SQLEXPRESS;Initial Catalog=master;
		    User ID=sa;Password=liebert;"; 
			OpenConnection(conn);
			try
			{
				ExcuteBakupQuery(conn);
			}
			catch(Exception ex)
			{
				Console.WriteLine("Error when excute backup query!"+ex.StackTrace.ToString());
			}
			CloseConnection(conn);		
			Console.WriteLine("****Finish to back up SQL Server Database*****");
		}
		
		/// <summary>
		/// This method is used to back up the database.
		/// Two type of database need to consider to be backup.
		/// </summary>
		public void BackUpDataBase()
		{
			switch(DB_DbType)
        	{
        		case 1:
					BackUpBundledDataBase();
        			break;
        		case 2:
        			BackUpSQLServerDataBase();
        			break;
        		default:
        			break;
        	}		
		}
		
		/// <summary>
		/// Restore bundled database;
		/// </summary>
		public void RestoreBundledDataBase()
		{
		  bool result = false;
		  Console.WriteLine("*****Start to restore bundled Database*****");     
		  string sourceDBPath = AppConfigOper.getConfigValue("DB_Bundled_Backup_Path");
		  string targetDBPath = AppConfigOper.getConfigValue("DB_Bundled_Path");
          
          //First delete the exsited directory.
		  if (!Directory.Exists(targetDBPath))
	      {
	       	  //Then create new directory with the same name.
	       	   Directory.CreateDirectory(targetDBPath);
	      }   
	      //Restore the files.
	      CopyDir(sourceDBPath,targetDBPath);
	      result = true;
	       Console.WriteLine("*****Finish to restore bundled Database*****");
	       RestoreResult = result;
		}
		
		/// <summary>
		/// Excute Resotre database query.
		/// </summary>
		/// <param name="conn">conn</param>
		public void ExcuteRestoreQuery(SqlConnection conn)
        {	
			bool result = false;
			string backupPath = @"C:\Nform\SqlBackup";
			string Nformbackup = @"C:\Nform\SqlBackup\Nform.bak";
			string NformAlmbackup = @"C:\Nform\SqlBackup\NformAlm.bak";
			string NformLogbackup = @"C:\Nform\SqlBackup\NformLog.bak";
			string movetoPath = @"C:\Nform\Moveto";
			if (!System.IO.Directory.Exists(backupPath))
            {
			//	MessageBox.Show("back up path is not existed!");
				return;
            }
			if(!File.Exists(Nformbackup))
			{
			//	MessageBox.Show("Nform back up file is not existed!");
				return;
			}
			if(!File.Exists(NformAlmbackup))
			{
			//	MessageBox.Show("NformAlm back up file is not existed!");
				return;
			}
			if(!File.Exists(NformLogbackup))
			{
			//	MessageBox.Show("NformLog back up file is not existed!");
				return;
			}
			if (!System.IO.Directory.Exists(movetoPath))
            {
				System.IO.Directory.CreateDirectory(movetoPath);
            }
			if (conn.State == ConnectionState.Closed) 
				conn.Open();
	        // Drop the three databases.
			string dropDB = @"drop database NformAlm;";
	        SqlCommand cmd = new SqlCommand(dropDB, conn);
	        cmd.ExecuteNonQuery();
	            
	        dropDB = @"drop database NformLog;";
	        cmd.CommandText = dropDB;
	        cmd.ExecuteNonQuery();
	                       
	        dropDB = @"drop database Nform;";
	        cmd.CommandText = dropDB;
	        cmd.ExecuteNonQuery();
	           
	        // Create 3 new database: Nform, NformAlm, NformLog.
	        string createDB=@"create database Nform;create database NformAlm;create database NformLog;";
			cmd.CommandText = createDB;
			cmd.ExecuteNonQuery();  
			    
			//Restore the existed database from backup files.
	        string restoreDB = @"restore database Nform from DISK = 'C:\Nform\SqlBackup\Nform.bak' with replace,
            MOVE 'Nform_data' TO 'C:\Nform\Moveto\data1.mdf',
            MOVE 'Nform_log' TO 'C:\Nform\Moveto\log1.ldf';";
		    cmd.CommandText = restoreDB;
		    cmd.ExecuteNonQuery();
          
	        restoreDB =  @"restore database NformLog from DISK = 'C:\Nform\SqlBackup\NformLog.bak' with replace,
            MOVE 'Nform_data' TO 'C:\Nform\Moveto\data2.mdf',
            MOVE 'Nform_log' TO 'C:\Nform\Moveto\log2.ldf';";
	        cmd.CommandText = restoreDB;
            cmd.ExecuteNonQuery();
	          
	        restoreDB =   @"restore database NformAlm from DISK = 'C:\Nform\SqlBackup\NformAlm.bak' with replace,
            MOVE 'Nform_data' TO 'C:\Nform\Moveto\data3.mdf',
            MOVE 'Nform_log' TO 'C:\Nform\Moveto\log3.ldf';";
	        cmd.CommandText = restoreDB; 
	        cmd.ExecuteNonQuery();
	        result = true;
	        
	        RestoreResult = result;
       }
		
		/// <summary>
		/// Restore SQL Server database
		/// </summary>
		public void RestoreSQLServerDataBase()
		{
			Console.WriteLine("*****Start to restore SQL Server Database*****");	        
	        SqlConnection conn = new SqlConnection();
	        conn.ConnectionString = GetDBConnString();
			OpenConnection(conn);
			try
			{
				ExcuteRestoreQuery(conn);
			}
			catch(Exception ex)
			{
				Console.WriteLine("Error when excute restore query!"+ex.StackTrace.ToString());
			}
			
			CloseConnection(conn);
			Console.WriteLine("*****Finish to restore SQL Server Database*****");
		}
		
		/// <summary>
		/// Restore the database, consider the type of database.
		/// DbType = 1, bundled database;DbType = 2, SQL Server database;
		/// </summary>
		public void RestoreDataBase()
		{
			if(BackUpResult == false)    //If Backup database is failed, restore database can not be excuted.
			{
				Console.WriteLine("Can not restore DB because of failure of backup database.");
				return;
			}
			else
			{
				switch(DB_DbType)
				{
	        		case 1:
						RestoreBundledDataBase();
	        			break;
	        		case 2:
	        			RestoreSQLServerDataBase();
	        			break;
	        		default:
	        		//	MessageBox.Show("Wrong database type!");
	        			break;
				}
			}
		}
		
		/// <summary>
		/// Delete Log file to clean the Nform.
		/// </summary>
		public void DeleteLogFile(String LogPath)
		{ 
		    if (File.Exists(LogPath))
		 {
			 System.IO.File.Delete(LogPath);
			 Console.WriteLine("*****Finish to Delete the Log File*****");
		 }	
		 else
		 	Console.WriteLine("*****This Log File is not existed.*****");
		}
		
		
				/// <summary>
		/// 
		/// </summary>
		/// <param name="ds"></param>
		/// <returns></returns>
		public static List<string[]> DataSetToList(DataSet ds)
		{
			List<string[]> strData = new List<string[]>();
			foreach(System.Data.DataRow r in ds.Tables[0].Rows)
			{
				int colCount = ds.Tables[0].Columns.Count;
				string[] items = new String[colCount];
				for(int i=0;i<colCount;i++)
				{
					items[i] = Convert.ToString(r.ItemArray[i]);
					string colName = ds.Tables[0].Columns[i].ColumnName;
					Console.WriteLine(colName +" : "+items[i]);
				}
				strData.Add(items);
			}
			return strData;
		}
		
		
	}
}
