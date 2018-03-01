using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace LogHelper
{
    public class LogHelper
    {

        private static LogHelper logHelper=null;

        private static readonly object lockobj = new object();

        LogHelper() { 
        
        
        }


        public static LogHelper getInstance()
        {
            lock (lockobj)
            {
                if (logHelper == null)
                    logHelper = new LogHelper();
            }
            return logHelper;

        }

        /// <summary>
        /// write log to special path log.txt ,second parameter is content
        /// </summary>
        /// <param name="FilePath"></param>
        /// <param name="LogStr"></param>
        public void LogIntoLocalFile(String FilePath,string LogStr) {
            Boolean floder_check = false;
            Boolean File_check = false;

            String Folder = getFolderath(FilePath);

            floder_check = CheckFolderExist(Folder);

            File_check = CheckFileExist(FilePath);

            try
            {
                if (floder_check && File_check)
                {
                    write2Txt(FilePath, LogStr);
                }
            }
            catch (Exception ee)
            {
                
                throw ee;
            }
        }

        public string getFolderath(String FilePath)
        {

            String[] path_arr= FilePath.Split('\\');

            string floder="";

            for (int i = 0; i < path_arr.Length-1;i++ )
            {
                floder = floder + path_arr[i] + "\\";
            
            }

            return floder;
        }


        public  Boolean CheckFileExist(String FilePath)
        {
            try
            {
                if (!File.Exists(FilePath))
                {
                    File.Create(FilePath).Dispose();
                    
                }
            }
            catch (Exception ee)
            {

                return false;
            }
            
            return true;
        }

        public  Boolean CheckFolderExist(String FolderPath)
        {
            try
            {
                if (!Directory.Exists(FolderPath))
                {
                    Directory.CreateDirectory(FolderPath);

                }
            }
            catch (Exception ee)
            {

                return false;
            }
            return true;
        } 

        public void write2Txt(String FilePath, string LogStr)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(FilePath, true))
            {
                //file.Write(LogStr);//直接追加文件末尾，不换行
                file.WriteLine(LogStr);// 直接追加文件末尾，换行 
                file.Close();
                file.Dispose();                    
            }
            
            //using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\testDir\test2.txt", true))
            //{
            //    foreach (string line in lines)
            //    {
            //        if (!line.Contains("second"))
            //        {
            //            file.Write(line);//直接追加文件末尾，不换行
            //            file.WriteLine(line);// 直接追加文件末尾，换行 
            //        }
            //    }
            //}

        }
        
    }
}
