using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Sql;
using System.Configuration;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Collections;
using System.Globalization;



//NPOI
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;


//Log

using LogHelper;

namespace ITVolume
{
    public partial class ChargeByStep : Form
    {

        
        public static string _configTest = System.Configuration.ConfigurationSettings.AppSettings["TestConStr"];

        public static string _configBA = System.Configuration.ConfigurationSettings.AppSettings["BAConStr"];

        public static string _configOSMQ = System.Configuration.ConfigurationSettings.AppSettings["OSMQConStr"];


        public static string app_configTest = System.Configuration.ConfigurationManager.ConnectionStrings["TestConStr"].ConnectionString;

        public static string app_configBA = System.Configuration.ConfigurationManager.ConnectionStrings["BAConStr"].ConnectionString;

        public static string app_configOSMQ = System.Configuration.ConfigurationManager.ConnectionStrings["OSMQConStr"].ConnectionString;

        static DateTimeFormatInfo dtFormat  = new System.Globalization.DateTimeFormatInfo();
        /*
         change db config here
         */
        //protected  const string TestDBstr = "Data Source=prcsgi10421d;Initial Catalog=BA;User Id=sa;Password=Infosys123";

        //protected const string OSMQDBstr = "Data Source=cnpek01808;Initial Catalog=OSMQ;User Id=wbofuser;Password=Wbgsnwbd1!";

        //protected const string BADBstr = "Data Source=cnpek0133d;Initial Catalog=BA;User Id=wbofuser;Password=Wbgsnwbd1!";

        protected string DefaultExcelPath = "D:\\ChargeData\\";

        public const string str_BAstastic = "sp_ba_n_step_";

        protected const string logFile = "\\log.txt";

        protected string directline_dir = "\\Importdata\\directline";

        protected string CDReport_dir = "\\Importdata\\CDReport";

        protected string snx_dir = "\\Importdata\\snx";

        protected string matrix_dir = "\\Importdata\\matrix";

        protected string cit_gid_dir = "\\Importdata\\cit_gid";

        protected string cit_email_dir = "\\Importdata\\cit_email";

        protected string cit_hr_dir = "\\Importdata\\hr";

        protected string cit_flender_dir = "\\Importdata\\cit_flender";

        protected string check_CC = "\\checkCC\\";

        protected string check_BU = "\\checkBU\\";

        protected string directline_dir_out = "\\outputdata\\directline";

        protected string snx_dir_out = "\\outputdata\\snx";

        protected string cit_hr_dir_out = "\\outputdata\\hr\\";

        protected string Imported_moved = "\\Imported\\";

        protected string BatchQ_config = "D:\\batchQ\\batchq.exe  -s cnpek01808:9170 -u BatchqUser -p Wbgsnwbd1! -c ";

        protected string ImportedFileKey="MovedFolderPath";

        //protected string ImportedFileFolderPath="";


        private Hashtable FileTable = null;

        private FileStream fs = null;

        private DataTable _dt = null;

        //private DataSet SheetSet = null;

        private String[] _SheetName_Arr = null;

        private String[] _ExcelName_Arr = null;

        private String _ExcelName = "", _SheetName = "";



        private String InvalidCCColNumName = "CHECKCC";

        
    static ChargeByStep(){
        dtFormat.ShortDatePattern = "yyyy-MM-dd";
    }

        public ChargeByStep()
        {
            InitializeComponent();

            ParaTime.Text = getChargeMonthNow();
            ExcelStorePath.Text = DefaultExcelPath;
            if (Initfloder())
                LogHelper.LogHelper.getInstance().LogIntoLocalFile(DefaultExcelPath + ParaTime.Text + logFile, "Init floder successfully");
               // MessageBox.Show("Init floder successfully");

            FileTable=InitFileHashtable(this.FileTable);

            if (LoaclDB.Checked)
                _configBA = _configTest;

            

            
        }

        private void ExcelStorePath_TextChanged(object sender, EventArgs e)
        {
            //DefaultExcelPath = ExcelStorePath.Text.Trim();

            //Initfloder();
            //FileTable = InitFileHashtable(this.FileTable);
        }

        private string getChargeMonthNow()
        {
            //return this.ParaTime.Text.Trim();

           return this.ParaTime.Text.Trim().Equals("")? DateTime.Now.ToString("yyyy-MM") + "-01":this.ParaTime.Text.Trim();



            //return DateTime.Now.ToString("yyyy-MM") + "-01";

        }




        public Boolean checkFolderExist(string path)
        {

            try
            {
                if (!Directory.Exists(path))
                {

                    Directory.CreateDirectory(path);
                   
                }
            }
            catch (Exception)
            {

                return false;
            }
            return true;
        }

        /// <summary>
        /// get osmq sp name
        /// </summary>
        /// <param name="relate"></param>
        /// <returns></returns>
        protected String GetOSMQSPName(String relate)
        {
            //int i = 1;

            relate = relate.Replace(" ", "_");

            string init = "sp_osmq_n_step_03_";

            return init + relate;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="StaticStr">static part like sp_ba_n_step_14_</param>
        /// <param name="relate">like 5_Volume</param>
        /// <returns>full name like sp_ba_n_step_14_5_Volume</returns>
        protected String GetRegularReportSpName(String StaticStr, String relate)
        {

            return StaticStr + relate;
        }

        protected IWorkbook GetWorkBook(string Excelname)
        {
            IWorkbook myworkbook = null;

            if (Excelname.IndexOf(".xlsx") > 0) // 2007版本
                myworkbook = new XSSFWorkbook();
            else if (Excelname.IndexOf(".xls") > 0) // 2003版本
                myworkbook = new HSSFWorkbook();
            else
            {
                msg.Text += ("invalid " + Excelname + " file");
                LogHelper.LogHelper.getInstance().LogIntoLocalFile(DefaultExcelPath + ParaTime.Text + logFile, "invalid " + Excelname + " file");

                return null;
            }
            msg.Text += ("create " + Excelname + " successfully");
            LogHelper.LogHelper.getInstance().LogIntoLocalFile(DefaultExcelPath + ParaTime.Text + logFile, "create " + Excelname + " successfully");
            return myworkbook;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fullExcelPath">excel full path need to create</param>
        /// <param name="workbook">abstract excel(workbook) object in memory</param>
        /// /// <param name="isEnd">whether excel is end when many sheets</param>
        /// <returns>{"success":true}</returns>
        private Boolean writeWorkBook2Excel(String fullExcelPath, IWorkbook workbook, Boolean isEnd)
        {

            string[] excel_name_arr = fullExcelPath.Split('\\');

            string excel_name = excel_name_arr[(excel_name_arr.Length - 1)];


            if ((!File.Exists(fullExcelPath)) && isEnd)
            {

                fs = new FileStream(fullExcelPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);

                try
                {
                    workbook.Write(fs);
                }
                catch (Exception ee)
                {

                    msg.Text += (excel_name + " meet issue" + "\r\n");
                    LogHelper.LogHelper.getInstance().LogIntoLocalFile(DefaultExcelPath + ParaTime.Text + logFile, excel_name + " meet issue; " + "exception:" + ee.ToString());
                    return false;
                }
                finally
                {
                    fs.Close();
                }

            }
            else if (File.Exists(fullExcelPath))
            {
                msg.Text += (excel_name + " have existed" + "\r\n");
                LogHelper.LogHelper.getInstance().LogIntoLocalFile(DefaultExcelPath + ParaTime.Text + logFile, excel_name + " have existed");
                return true;

            }
            return true;
        }


        /// <summary>
        /// write datatable data to named excel(workbook) sheet
        /// </summary>
        /// <param name="workbook">excel interface ref need to write</param>
        /// <param name="dt">datatable need to write as datasource</param>
        /// <param name="SheetName">be wirtten excel sheet name store data</param>
        /// <param name="HaveHeader">true:write excel head in first column</param>
        /// <param name="HaveHeader">true:write excel head in first column</param>
        /// <returns>bool result,true means success,fail means exception</returns>
        private Boolean CopyDT2ExcelSheet(IWorkbook workbook, DataTable dt, String SheetName, Boolean HaveHeader, out IWorkbook outworkbook)
        {
            ISheet isheet;
            IRow irow;
            ICellStyle style;
            IFont f;

            int count = 0, row = 0, column, i, j;

            if (workbook == null)
            {
                outworkbook = workbook;
                return false;
            }

            isheet = workbook.CreateSheet(SheetName);

            if (HaveHeader)
            {
                irow = isheet.CreateRow(0);

                /*               
                 set column header font blod                 
                 */
                style = workbook.CreateCellStyle();
                f = workbook.CreateFont();
                f.Boldweight = (short)FontBoldWeight.Bold;
                style.SetFont(f);

                for (i = 0; i < dt.Columns.Count; i++)
                {
                    irow.CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);//get column name as first row
                    irow.GetCell(i).CellStyle = style;
                }
                count = 1;
            }
            row = dt.Rows.Count;
            column = dt.Columns.Count;
            for (j = 0; (count + j) < (row + count); j++)
            {
                irow = isheet.CreateRow(count + j);//create row

                for (i = 0; i < column; i++)

                    irow.CreateCell(i).SetCellValue((dt.Rows[j][i]).ToString());
                //copy dt data here
            }
            outworkbook = workbook;
            return true;
        }

        /// <summary>
        /// 执行存储过程，返回datatable
        /// 
        /// </summary>
        /// <param name="spname"></param>
        /// <param name="Paraname">不需要参数就直接传""</param>
        /// <param name="Paravalue"></param>
        /// <param name="constr"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        protected DataTable ExcuteQuerySP(String spname, String Paraname, String Paravalue, String constr)
        {
            DataTable dt = new DataTable();
            //message = "success";
            SqlConnection sqlcon;
            SqlDataAdapter da;
            SqlParameter sqlpara;
            try
            {
                sqlcon = new SqlConnection(constr);
            }
            catch (Exception ee)
            {
                LogHelper.LogHelper.getInstance().LogIntoLocalFile(DefaultExcelPath + ParaTime.Text, ee.ToString());
                return dt;
            }
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = spname;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = sqlcon;

            if (Paraname != "")
            {
                sqlpara = new SqlParameter(Paraname, Paravalue);
                cmd.Parameters.Add(sqlpara);
            }
            else if (Paraname == "")
            {

                sqlpara = new SqlParameter();
            }
            try
            {
                sqlcon.Open();
            }
            catch (Exception ee)
            {
                LogHelper.LogHelper.getInstance().LogIntoLocalFile(DefaultExcelPath + ParaTime.Text, ee.ToString());
                return dt;
            }
            da = new SqlDataAdapter();
            da.SelectCommand = cmd;
            da.Fill(dt);

            sqlcon.Close();

            return dt;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="StrSQL"></param>
        /// <param name="constr"></param>
        /// <returns></returns>
        protected DataTable ExcuteQuerySQL(String StrSQL,  String constr)
        {
            DataTable dt = new DataTable();
            //message = "success";
            SqlConnection sqlcon;
            SqlDataAdapter da;
            //SqlParameter sqlpara;
            try
            {
                sqlcon = new SqlConnection(constr);
            }
            catch (Exception ee)
            {
                LogHelper.LogHelper.getInstance().LogIntoLocalFile(DefaultExcelPath + ParaTime.Text, ee.ToString());
                return dt;
            }
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = StrSQL;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlcon;

           
            try
            {
                sqlcon.Open();
            }
            catch (Exception ee)
            {
                LogHelper.LogHelper.getInstance().LogIntoLocalFile(DefaultExcelPath + ParaTime.Text, ee.ToString());
                return dt;
            }
            da = new SqlDataAdapter();
            da.SelectCommand = cmd;
            da.Fill(dt);

            sqlcon.Close();

            return dt;
        }

        /// <summary>
        /// 执行存储过程，返回datatable
        /// </summary>
        /// <param name="spname"></param>
        /// <param name="Paraname"></param>
        /// <param name="Paravalue"></param>
        /// <param name="constr"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        protected DataTable ExcuteQuerySPOutFlag(String spname, String Paraname, String Paravalue, String constr,string intSP_OutParaname ,out int intflag)
        {
            DataTable dt = new DataTable();
            //message = "success";
            SqlConnection sqlcon;    
            SqlDataAdapter da;
            SqlParameter sqlpara;
            try
            {
                sqlcon = new SqlConnection(constr);
            }
            catch (Exception ee)
            {
                intflag = 0;
                return dt;
            }
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = spname;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = sqlcon;

            if (Paraname != "")
            {
                sqlpara = new SqlParameter(Paraname, Paravalue);
                
                cmd.Parameters.Add(sqlpara);
            }
            else if (Paraname == "")
            {

                sqlpara = new SqlParameter();
            }

            sqlpara = new SqlParameter("@flag",SqlDbType.Int);
            cmd.Parameters.Add(sqlpara);
            cmd.Parameters["@flag"].Direction = System.Data.ParameterDirection.Output; 
            try
            {
                sqlcon.Open();
            }
            catch (Exception ee)
            {
                intflag = 0;
                LogHelper.LogHelper.getInstance().LogIntoLocalFile(DefaultExcelPath + ParaTime.Text, ee.ToString());
                return dt;
            }
            da = new SqlDataAdapter();
            da.SelectCommand = cmd;        
            da.Fill(dt);
            intflag = (int)cmd.Parameters[intSP_OutParaname].Value;

            sqlcon.Close();

            return dt;
        }


        private String get_XLS_ExcelName(string folderpath, string excelname)
        {

            return folderpath + "\\" + excelname + ".xls";
        }

        private String get_XLSX_ExcelName(string folderpath, string excelname)
        {

            return folderpath + "\\" + excelname + ".xlsx";
        }

        /// <summary>
        /// create excel multiple sheets or single sheet
        /// </summary>
        /// <param name="FolderPath">excel floder</param>
        /// <param name="ExcelName">excel name</param>
        /// <param name="sheet">sheet name</param>
        /// <param name="dt">datatable copy to sheet</param>
        /// <param name="Myworkbook">abstract excel(workbook) ref</param>
        /// <param name="msg">form text object ref</param>
        /// <param name="isEnd">mark as flase is multiple</param>
        /// <param name="outMyworkbook">out workbook ref</param>
        /// <returns></returns>
        private Boolean CommonExecute(String FolderPath, String ExcelName, String sheet, DataTable dt, IWorkbook Myworkbook, TextBox msg, Boolean isEnd, out IWorkbook outMyworkbook)
        {
            Boolean Copydata2Sheet = false;
            try
            {
                if (!checkFolderExist(FolderPath))
                {
                    msg.Text += "Create default path fail in " + FolderPath;
                    LogHelper.LogHelper.getInstance().LogIntoLocalFile(DefaultExcelPath + ParaTime.Text + logFile, "Create default path fail in " + FolderPath);
                }
                if (Myworkbook == null)
                    Myworkbook = GetWorkBook(get_XLS_ExcelName(FolderPath, ExcelName));

                Copydata2Sheet = CopyDT2ExcelSheet(Myworkbook, dt, sheet, true, out Myworkbook);

                if (Copydata2Sheet)
                {
                    writeWorkBook2Excel(get_XLS_ExcelName(FolderPath, ExcelName), Myworkbook, isEnd);
                }

            }
            catch (Exception ee)
            {
                msg.Text = "issue occurs in " + ExcelName + ",sheet" + sheet + "\r\n";
                LogHelper.LogHelper.getInstance().LogIntoLocalFile(DefaultExcelPath + ParaTime.Text + logFile, "issue occurs in " + ExcelName + ", sheet " + sheet+" ;exception:"+ee.ToString());
                outMyworkbook = null;
                return false;
            }
            outMyworkbook = Myworkbook;
            return true;
        }


        private void Test_Click(object sender, EventArgs e)
        {
            _SheetName_Arr = new String[] { "test1", "test2" };
            _ExcelName = "test";
            IWorkbook Myworkbook = null;
            Boolean isEnd = false;

            int count_test = 0;

            foreach (string sheetname in _SheetName_Arr)
            {
                count_test++;
                if (count_test == 2)
                {
                    isEnd = true;
                }

                DataTable dt = ExcuteQuerySP(GetOSMQSPName(_ExcelName), "", "", app_configTest);
                dt.TableName = sheetname;
                CommonExecute(DefaultExcelPath + ParaTime.Text + "\\report", _ExcelName, sheetname, dt, Myworkbook, msg, isEnd, out Myworkbook);
            }

            LogHelper.LogHelper.getInstance().LogIntoLocalFile(DefaultExcelPath + ParaTime.Text + "\\log.txt", "test");
        }

        private void ITVolume_Click(object sender, EventArgs e)
        {
            //_SheetName_Arr = new String[] { "test" };
            IWorkbook Myworkbook = null;
            _ExcelName_Arr = new String[] { "Database Hosting", "Disk Space", "Home Drive", "Public Folder", "SharePoint DMS(Collaboration)", "Telephone Conference", "Unlimited E-Mail", "WCMS", "Web Hosting" };


            foreach (string ExcelName in _ExcelName_Arr)
            {
                _SheetName = ExcelName;
                DataTable dt = ExcuteQuerySP(GetOSMQSPName(ExcelName), "", "", app_configOSMQ);
                dt.TableName = _SheetName;
                CommonExecute(DefaultExcelPath + ParaTime.Text + "\\report", ExcelName, _SheetName, dt, Myworkbook, msg, true, out Myworkbook);
                Myworkbook = null;
            }
        }


        

        private void Voice_fault_data_Click(object sender, EventArgs e)
        {
            _SheetName_Arr = new String[] { "G4", "Mobile", "Double_G4", "Invalid_Share_CC", "Double_Voice", "Inactive_Telephone" };
            _ExcelName = "Vioce Fault Data_" + ParaTime.Text.Trim();

            IWorkbook Myworkbook = null;
            Boolean isEnd = false;
            string sheetname = "";
            for (int i = 0; i < _SheetName_Arr.Length; i++)
            {
                sheetname = _SheetName_Arr[i];
                if (i == _SheetName_Arr.Length - 1)
                    isEnd = true;

                DataTable dt = ExcuteQuerySP(GetRegularReportSpName("sp_ba_n_step_30_", sheetname), "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                dt.TableName = sheetname;

                if (_SheetName_Arr[i].Length < 13)
                    sheetname = sheetname + "_Volume";

                sheetname = sheetname.Replace('_', ' ');

                CommonExecute(DefaultExcelPath + ParaTime.Text + "\\report", _ExcelName, sheetname, dt, Myworkbook, msg, isEnd, out Myworkbook);
            }
        }

        private void double_charge_Click(object sender, EventArgs e)
        {
            _SheetName_Arr = new String[] { "Onetime", "Volume", "Monthly", "Special", "Invalid Special" };
            _ExcelName = "Double Charge_" + ParaTime.Text.Trim();

            IWorkbook Myworkbook = null;
            Boolean isEnd = false;
            string sheetname = "";
            for (int i = 0; i < _SheetName_Arr.Length; i++)
            {
                sheetname = _SheetName_Arr[i];
                if (i == _SheetName_Arr.Length - 1)
                    isEnd = true;

                DataTable dt = ExcuteQuerySP(GetRegularReportSpName("sp_ba_n_step_35_", sheetname), "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                dt.TableName = sheetname;


                CommonExecute(DefaultExcelPath + ParaTime.Text + "\\report", _ExcelName, sheetname, dt, Myworkbook, msg, isEnd, out Myworkbook);
            }
        }

        private void atos_Click(object sender, EventArgs e)
        {
            //String SPName = "";

            String ExcelName = "";
            IWorkbook Myworkbook = null;
            Boolean isEnd = true;
            DataTable dt = null;
            string sheetname = "sheet1";


            ExcelName = "Atos_costcenter";
            dt = ExcuteQuerySP(getAtosSPName(ExcelName), "@InputMonth", ParaTime.Text.Trim(), app_configBA);
            CommonExecute(DefaultExcelPath + ParaTime.Text + "\\report", ExcelName, sheetname, dt, Myworkbook, msg, isEnd, out Myworkbook);
            if (Myworkbook != null)
                Myworkbook = null;


            ExcelName = "5_costcenter";
            dt = ExcuteQuerySP(getAtosSPName(ExcelName), "@InputMonth", ParaTime.Text.Trim(), app_configBA);
            CommonExecute(DefaultExcelPath + ParaTime.Text + "\\report", ExcelName, sheetname, dt, Myworkbook, msg, isEnd, out Myworkbook);
            if (Myworkbook != null)
                Myworkbook = null;

            ExcelName = "27_Costcenter";
            dt = ExcuteQuerySP(getAtosSPName(ExcelName), "@InputMonth", ParaTime.Text.Trim(), app_configBA);
            CommonExecute(DefaultExcelPath + ParaTime.Text + "\\report", ExcelName, sheetname, dt, Myworkbook, msg, isEnd, out Myworkbook);
            if (Myworkbook != null)
                Myworkbook = null;

            ExcelName = "154_costcenters";
            dt = ExcuteQuerySP(getAtosSPName(ExcelName), "@InputMonth", ParaTime.Text.Trim(), app_configBA);
            CommonExecute(DefaultExcelPath + ParaTime.Text + "\\report", ExcelName, sheetname, dt, Myworkbook, msg, isEnd, out Myworkbook);
            if (Myworkbook != null)
                Myworkbook = null;

        }

        protected String getAtosSPName(String related)
        {
            return "sp_ba_n_step_43_" + related + "_select";
        }



        public Boolean TravelDirectLineExcelInFloder(String floder)
        {

            IWorkbook myworkbook = null;

            SqlBulkCopy sqlbulkcopy = null;

            string[] Dir = System.IO.Directory.GetFiles(floder);

            DataTable dt = new DataTable();

            dt = initDirectLine_Datatable(dt);

            foreach (string file in Dir)// file stand full path
            {
                SetFullFileNameIntoHashtable(FileTable, file);
                System.IO.FileInfo FI = new System.IO.FileInfo(file);
                if ((FI.Extension != ".xls") && (FI.Extension != ".xlsx"))
                {
                    return false;
                }
                myworkbook = GetworkBookwithDatafromLocal(file);

                dt = PrepareDirectlineDataTable(myworkbook, dt);
                //change db config here
                sqlbulkcopy = initDirectLine_SqlBulkCopy(sqlbulkcopy, _configBA ,"tblvoicevolume");

                DataTable2DB(sqlbulkcopy, dt);

                dt.Clear();

                MoveImportedFile(file);

                LogHelper.LogHelper.getInstance().LogIntoLocalFile(DefaultExcelPath + ParaTime.Text, file);
            }

            return true;
        }




        public Boolean TravelSnxExcelInFloder(String floder){

            Boolean snx = false,bcp = false;

            string[] Dir = System.IO.Directory.GetFiles(floder);       

            foreach (string file in Dir)// file stand full path
            {
                if(file.ToUpper().Contains("SNX_")){
                    snx=ImportSnxExcel(file);
                    
                    if (snx)
                        MoveImportedFile(file);
                }

                else if (file.ToUpper().Contains("BCP_"))
                {                    
                    bcp=ImportBcpExcel(file);

                    if (bcp)
                        MoveImportedFile(file);
                }

                else {
                    return false;              
                }
            }

            return (snx && bcp);
        
        }

        public Boolean TravelMatrixExcelInFloder(String floder)
        {

            Boolean matrix = false;

            string[] Dir = System.IO.Directory.GetFiles(floder);

            foreach (string file in Dir)// file stand full path
            {
                if (file.ToUpper().Contains("MATRIX"))
                {
                    matrix = ImportMatrixExcel(file);
                }

                else
                {
                    return false;
                }
            }

            return (matrix);

        }


        public Boolean TravelCIT_GIDExcelInFloder(String floder)
        {

            Boolean CIT_GID = false;

            int true_count=0;

            string[] Dir = System.IO.Directory.GetFiles(floder);

            foreach (string file in Dir)// file stand full path
            {
                SetFullFileNameIntoHashtable(FileTable, file);

                CIT_GID = ImportCIT_GIDExcel(file);

                if (CIT_GID)
                {
                    true_count++;
                    MoveImportedFile(file);
                }

            }

            if (true_count == Dir.Length)
                return true;
            
            return false;

        }

        public Boolean TravelCDReportExcelInFloder(String floder)
        {

            Boolean CDReport = false;

            int true_count = 0;

            string[] Dir = System.IO.Directory.GetFiles(floder);

            foreach (string file in Dir)// file stand full path
            {
                SetFullFileNameIntoHashtable(FileTable, file);

                CDReport = ImportCDReportExcel(file);

                if (CDReport)
                {
                    true_count++;
                    MoveImportedFile(file);
                }

            }

            if (true_count == Dir.Length)
                return true;

            return false;

        }



        public Boolean TravelCheckCCExcelInFloder(String floder)
        {

            Boolean CheckCC = false;

            int true_count = 0;

            string[] Dir = System.IO.Directory.GetFiles(floder);

            foreach (string file in Dir)// file stand full path
            {
                SetFullFileNameIntoHashtable(FileTable, file);

                CheckCC = ImportCheckCCExcelandOutPutInvalidCC(file);

                if (CheckCC)
                {
                    true_count++;
                    MoveImportedFile(file);
                }

            }

            if (true_count == Dir.Length)
                return true;

            return false;

        }

        public Boolean TravelCIT_EmailExcelInFloder(String floder)
        {

            Boolean CIT_Email = false;

            int true_count = 0;

            string[] Dir = System.IO.Directory.GetFiles(floder);

            foreach (string file in Dir)// file stand full path
            {
                SetFullFileNameIntoHashtable(FileTable, file);
                CIT_Email = ImportCIT_EmailExcel(file);

                if (CIT_Email)
                {
                    MoveImportedFile(file);
                    true_count++;
                }
            }

            if (true_count == Dir.Length)
                return true;

            return false;

        }

        public Boolean TravelCIT_FlenderExcelInFloder(String floder)
        {

            Boolean CIT_Flender = false;

            int true_count = 0;

            string[] Dir = System.IO.Directory.GetFiles(floder);

            foreach (string file in Dir)// file stand full path
            {
                SetFullFileNameIntoHashtable(FileTable, file);
                CIT_Flender = ImportCIT_FlenderExcel(file);

                if (CIT_Flender)
                {
                    true_count++;
                    MoveImportedFile(file);
                }
            }

            if (true_count == Dir.Length)
                return true;

            return false;

        }



        public Boolean ImportCIT_EmailExcel(String filepath)
        {

            IWorkbook myworkbook = null;

            SqlBulkCopy sqlBulkCopy = null;

            DataTable dt = new DataTable();

            dt = initCIT_Email_Datatable(dt);

            myworkbook = GetworkBookwithDatafromLocal(filepath);

            PrepareCIT_EmailDataTable(myworkbook, dt);

            sqlBulkCopy = initCIT_Email_SqlBulkCopy(sqlBulkCopy, _configBA, "tmpEmail2");

            return DataTable2DB(sqlBulkCopy, dt);


        }

        public Boolean ImportCIT_FlenderExcel(String filepath)
        {

            IWorkbook myworkbook = null;

            SqlBulkCopy sqlBulkCopy = null;

            DataTable dt = new DataTable();

            dt = initCIT_Flender_Datatable(dt);

            myworkbook = GetworkBookwithDatafromLocal(filepath);

            PrepareCIT_FlenderDataTable(myworkbook, dt);

            sqlBulkCopy = initCIT_Flender_SqlBulkCopy(sqlBulkCopy, _configBA, "tbltmpcharges_tmp");

            return DataTable2DB(sqlBulkCopy, dt);

        }
      
        public Boolean ImportCIT_GIDExcel(String filepath)
        {

            IWorkbook myworkbook = null;

            SqlBulkCopy sqlBulkCopy = null;

            DataTable dt = new DataTable();

            dt = initCIT_GID_Datatable(dt);

            myworkbook = GetworkBookwithDatafromLocal(filepath);

            PrepareCIT_GIDDataTable(myworkbook, dt);

            sqlBulkCopy = initCIT_GID_SqlBulkCopy(sqlBulkCopy, _configBA, "tblCITBASE1");

            return DataTable2DB(sqlBulkCopy, dt);

        }

        public Boolean ImportCDReportExcel(String filepath)
        {

            IWorkbook myworkbook = null;

            SqlBulkCopy sqlBulkCopy = null;

            String FileNameWithoutSuffix = "";

            DataTable dt = new DataTable();

            dt = initCDReport_Datatable(dt);

            myworkbook = GetworkBookwithDatafromLocal(filepath);

            FileNameWithoutSuffix = GetFileNameFromFullFileNameWithoutSuffix(filepath);

            PrepareCDReportDataTable(myworkbook, dt, FileNameWithoutSuffix);

            sqlBulkCopy = initCDReportSqlBulkCopy(sqlBulkCopy, _configBA, "tblg4volume");

            return DataTable2DB(sqlBulkCopy, dt);

        }

        public Boolean ImportCheckCCExcelandOutPutInvalidCC(String filepath)
        {
            Boolean Copydata2Sheet = false;

            IWorkbook myworkbook = null;
            //store excel data
            DataTable Exceldt = new DataTable();
            //store valid data
            DataTable Validdt = new DataTable();

            DataTable NewExceldt = null;

            DataRow[] Arr_InvalidRow = null;

            //string[] ColumnNameArr = null;

            //string _CheckCC_Column="CheckCC";

            string ValidSQL = "select distinct ccostcenter from vwtblcostcenter";

            string ValidCondition = "";

            string LeftValidCCSheetName = "#Valid#CC#";

            string InvalidCCSheetName = "#Invalid#CC#";

            //dt = initCIT_GID_Datatable(dt);
            
            /*
             init DT according to excel column
             */

            myworkbook = GetworkBookwithDatafromLocal(filepath); 

            //ColumnNameArr=GetFirstRowASColumnName(myworkbook.GetSheetAt(0).GetRow(0));

            Exceldt=DataTable1stLineAsDTNameFromExcel(myworkbook, Exceldt);

            //Exceldt = PrepareDataTableFromColumnName(ColumnNameArr, Exceldt);

            Validdt = ExcuteQuerySQL(ValidSQL, _configBA);

            ValidCondition = getValidCOL_LINQ(InvalidCCColNumName, Validdt);

            NewExceldt=Exceldt.Clone();

            Arr_InvalidRow = Exceldt.Select(ValidCondition);

            foreach (DataRow dr in Arr_InvalidRow)
            {
                NewExceldt.ImportRow(dr);

                Exceldt.Rows.Remove(dr);
            }

            filepath=ChangeFilePath(filepath, InvalidCCColNumName);


            /*Invalid Part*/
            Copydata2Sheet = CopyDT2ExcelSheet(myworkbook, NewExceldt, InvalidCCSheetName, true, out myworkbook);

            //if (Copydata2Sheet)
            //{
            //    writeWorkBook2Excel(filepath, myworkbook, true);
            //}

            /*Valid Part*/
            Copydata2Sheet = (Copydata2Sheet) && (CopyDT2ExcelSheet(myworkbook, Exceldt, LeftValidCCSheetName, true, out myworkbook));

            if (Copydata2Sheet)
            {
                writeWorkBook2Excel(filepath, myworkbook, true);
            }


             return true;
        }

        public DataTable DataTable1stLineAsDTNameFromExcel(IWorkbook workbook,DataTable dt) {
            String[] ColumnNameArr = null;

            int column = 0, rownum = 0;

            Boolean CheckCC = false;

            ISheet sheet = null;

            DataRow dr = null;
            //get sheet obj here
            sheet = workbook.GetSheetAt(0);

            ColumnNameArr = GetFirstRowASColumnName(sheet.GetRow(0));

            column = ColumnNameArr.Length-1;

            rownum = sheet.LastRowNum;    //LastRowNum from 0,
            //dt.Columns.Add("id", Type.GetType("System.Int32"));

          

            foreach (String colnum in ColumnNameArr)
            {
                dt.Columns.Add(colnum, Type.GetType("System.String"));

                if (colnum.Equals(InvalidCCColNumName))
                    CheckCC = true;
            }


            if (!CheckCC)
            {
                //dt.Clear();

                return null;
            }

            for (int i = 1; i <= rownum; i++)//notice num here,from 1 not read first row data
            {
                dr = dt.NewRow();
                for (int j = 0; j <= column; j++)
                {
                    
                    //check logic
                    CheckCellBlank(sheet.GetRow(i), j, CellType.String);
                    //must use i-1 to suit 0 from dt
                    dr[j] = sheet.GetRow(i).GetCell(j).ToString();

                    //dt.Rows.

                    
                }
                dt.Rows.Add(dr);
                dr = null;
            }


            return dt;


            //return null;
        }

        /// <summary>
        /// datatable contains valid conditon single column data
        /// checked colnum must use name special , like CHECKCC
        /// getValidCOL_LINQ("CHECKCC",dt)
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        public String getValidCOL_LINQ(String columnName,DataTable dt) {

            String Where = "";

            String VWCostCenterValue="ccostcenter";

            Where = Where + columnName + " not in ( ";

            int row = 0;
            row = dt.Rows.Count;

            for (int i = 0;i < row; i++)  
            {
                Where = Where + "'" + dt.Rows[i][VWCostCenterValue].ToString() + "'";
                if (i == row - 1)
                    break;
                Where = Where + ",";
            }

            Where = Where + "  )";

            return Where;
        }

        public string[] GetFirstRowASColumnName(IRow irow) {

            int col_num = irow.LastCellNum;

            string[] arr_column = new string[col_num];


            for (int i=0;i<col_num ;i++ ) {
                CheckCellBlank(irow, i, CellType.String);

                arr_column[i] = irow.GetCell(i).StringCellValue;
            
            }

            return arr_column;
        }

        public Boolean ImportMatrixExcel(String filepath)
        {

            IWorkbook myworkbook = null;

            SqlBulkCopy sqlBulkCopy = null;

            DataTable dt = new DataTable();

            SetFullFileNameIntoHashtable(FileTable, filepath);

            dt = initMatrix_Datatable(dt);

            myworkbook = GetworkBookwithDatafromLocal(filepath);

            PrepareMatrixDataTable(myworkbook, dt);

            sqlBulkCopy = initMatrix_SqlBulkCopy(sqlBulkCopy, _configBA, "tbltempMatrix");

          

            return DataTable2DB(sqlBulkCopy, dt);

            //SqlBulkCopy sqlbulkcopy = null;

            //DataTable dt = new DataTable();

            //dt = initDirectLine_Datatable(dt);

            // return true;
        }



        public Boolean ImportSnxExcel(String filepath) {

            IWorkbook myworkbook = null;

            SqlBulkCopy sqlBulkCopy=null;

            SetFullFileNameIntoHashtable(FileTable, filepath);

            DataTable dt = new DataTable();

            dt = initSNX_Datatable(dt);

            myworkbook = GetworkBookwithDatafromLocal(filepath);

            PrepareSNXDataTable(myworkbook, dt);

            sqlBulkCopy = initSnx_SqlBulkCopy(sqlBulkCopy, _configBA, "tbltempSCDinfo");

            

            return DataTable2DB(sqlBulkCopy, dt);

            //SqlBulkCopy sqlbulkcopy = null;

            //DataTable dt = new DataTable();

            //dt = initDirectLine_Datatable(dt);

           // return true;
        }


        public Boolean ImportBcpExcel(String filepath)
        {

            IWorkbook myworkbook = null;

            SqlBulkCopy sqlBulkCopy = null;

            SetFullFileNameIntoHashtable(FileTable, filepath);

            DataTable dt = new DataTable();

            dt = initBcp_Datatable(dt);

            myworkbook = GetworkBookwithDatafromLocal(filepath);

            PrepareBcpDataTable(myworkbook, dt);

            sqlBulkCopy = initBcp_SqlBulkCopy(sqlBulkCopy, _configBA, "tbltempSCD");

            

            return DataTable2DB(sqlBulkCopy, dt);
           
        }


        public Boolean ImportHrExcel(String filepath)
        {

            IWorkbook myworkbook = null;

            SqlBulkCopy sqlBulkCopy = null;

            SetFullFileNameIntoHashtable(FileTable, filepath);

            DataTable dt = new DataTable();

            dt = initHr_Datatable(dt);

            myworkbook = GetworkBookwithDatafromLocal(filepath);

            dt = PrepareHrDataTable(myworkbook, dt);

            sqlBulkCopy = initHr_SqlBulkCopy(sqlBulkCopy, _configBA, "TempHR");

            return DataTable2DB(sqlBulkCopy, dt);

        }

        /// <summary>
        /// create excel file and return abstract excel IWorkbook object 
        /// </summary>
        /// <param name="path">full excel file path</param>
        /// <returns></returns>
        protected IWorkbook GetworkBookwithDatafromLocal(string path)
        {
            IWorkbook myworkbook = null;

            fs = new FileStream(path, FileMode.Open, FileAccess.Read);
            if (path.ToLower().IndexOf(".xlsx") > 0) // 2007版本
                myworkbook = new XSSFWorkbook(fs);
            else if (path.ToLower().IndexOf(".xls") > 0) // 2003版本
                myworkbook = new HSSFWorkbook(fs);

            fs.Close();

            fs.Dispose();

            return myworkbook;

        }

        protected DataTable PrepareDirectlineDataTable(IWorkbook myworkbook, DataTable dt)
        {

            // 3 files sheet name are Sheet1
            ISheet sheet = myworkbook.GetSheet("Sheet1");

            DataRow DR = null;

            IRow irow = null;

            int colNum = sheet.GetRow(0).LastCellNum + 1;

            for (int i = 1; i < sheet.LastRowNum+1; i++)
            {
                DR = dt.NewRow();

                irow = sheet.GetRow(i);
                irow.GetCell(0).SetCellType(CellType.String);
                //irow.GetCell(1).SetCellType(CellType.Numeric);
                //irow.GetCell(2).SetCellType(CellType.Numeric);
                //if (irow.GetCell(0).CellType == CellType.Blank)

                //remove blank telephone  chinese char
                if ((irow.GetCell(0).StringCellValue.Trim() == "")||checkChineseStr(irow.GetCell(0).StringCellValue.Trim()))
                    continue;

                // DR["id"] = i;
                DR["dchargemonth"] = getChargeMonthNow();
                DR["telephone"] = irow.GetCell(0).StringCellValue;
                DR["starttime"] = irow.GetCell(1).DateCellValue.ToString("yyyy-MM-dd");
                DR["stoptime"] = irow.GetCell(2).DateCellValue.ToString("yyyy-MM-dd");
                DR["totalprice"] = irow.GetCell(3).NumericCellValue;
                DR["istatus"] = "0";
                DR["ceditby"] = "z003pkxs";
                DR["dedittime"] = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");

                dt.Rows.Add(DR);
            }

            return dt;
        }

        protected DataTable PrepareSNXDataTable(IWorkbook myworkbook, DataTable dt)
        {
            ISheet sheet = myworkbook.GetSheetAt(0);

            DataRow DR = null;

            IRow irow = null;

            int colNum = sheet.GetRow(0).LastCellNum + 1;

            for (int i = 1; i < sheet.LastRowNum + 1; i++)
            {
                DR = dt.NewRow();

                irow = sheet.GetRow(i);
                
                //if (irow.GetCell(27)==null)
                //{
                //    irow.CreateCell(27,CellType.String);
                //    //irow.GetCell(27) = irow.;

                //    irow.GetCell(27).SetCellValue("");
                    
                //}
                CheckCellBlank(irow, 35, CellType.String);
                CheckCellBlank(irow,27, CellType.String);
                CheckCellBlank(irow,26, CellType.String);

                irow.GetCell(26).SetCellType(CellType.String);
                irow.GetCell(27).SetCellType(CellType.String);
                irow.GetCell(4).SetCellType(CellType.String);
                irow.GetCell(3).SetCellType(CellType.String);
                irow.GetCell(1).SetCellType(CellType.String);
                //irow.GetCell(1).SetCellType(CellType.Numeric);
                //irow.GetCell(2).SetCellType(CellType.Numeric);
                //if (irow.GetCell(0).CellType == CellType.Blank)

                //remove blank telephone  chinese char
                //if ((irow.GetCell(4).StringCellValue.Trim() == "") || checkChineseStr(irow.GetCell(4).StringCellValue.Trim()))
                    //continue;
                // DR["id"] = i;
                DR["dchargemonth"] = getChargeMonthNow();
                DR["ccostcenter"] = irow.GetCell(26).StringCellValue;
                DR["cperno"] = irow.GetCell(27).StringCellValue;
                DR["cname"] = irow.GetCell(4).StringCellValue.Replace('\'',' ');//this column may contain "\'"
                DR["clocation"] = irow.GetCell(3).StringCellValue;
                DR["corganization"] = irow.GetCell(1).StringCellValue;
                DR["ceditby"] = "z003pkxs";
                DR["dedittime"] = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                DR["GID"] = irow.GetCell(35).StringCellValue; ;

                dt.Rows.Add(DR);
            }

            return dt;
        }


        protected DataTable PrepareBcpDataTable(IWorkbook myworkbook, DataTable dt)
        {
            ISheet sheet = myworkbook.GetSheetAt(0);

            DataRow DR = null;

            IRow irow = null;

            int colNum = sheet.GetRow(0).LastCellNum + 1;

            for (int i = 1; i < sheet.LastRowNum + 1; i++)
            {
                DR = dt.NewRow();

                irow = sheet.GetRow(i);

                CheckCellBlank(irow, 16, CellType.String);
                CheckCellBlank(irow, 26, CellType.String);
                CheckCellBlank(irow, 27, CellType.String);

                irow.GetCell(35).SetCellType(CellType.String);
                irow.GetCell(26).SetCellType(CellType.String);
                irow.GetCell(27).SetCellType(CellType.String);
                irow.GetCell(4).SetCellType(CellType.String);
                //if (irow.GetCell(16) == null)
                //{
                //    irow.CreateCell(16, CellType.String);
                //    //irow.GetCell(27) = irow.;

                //    irow.GetCell(16).SetCellValue("");

                //}
                //irow.GetCell(16).SetCellType(CellType.String);
                //irow.GetCell(1).SetCellType(CellType.String);
                //irow.GetCell(1).SetCellType(CellType.Numeric);
                //irow.GetCell(2).SetCellType(CellType.Numeric);
                //if (irow.GetCell(0).CellType == CellType.Blank)

                //remove blank telephone  chinese char
                //if ((irow.GetCell(4).StringCellValue.Trim() == "") || checkChineseStr(irow.GetCell(4).StringCellValue.Trim()))
                //continue;
                // DR["id"] = i;
                DR["dchargemonth"] = getChargeMonthNow();            
                DR["cperno"] = irow.GetCell(27).StringCellValue;
                DR["cname"] = irow.GetCell(4).StringCellValue.Replace('\'', ' ');//this column may contain "\'"
                DR["cemail"] = irow.GetCell(16).StringCellValue;
                DR["ccostcenter"] = irow.GetCell(26).StringCellValue;               
                DR["ceditby"] = "z003pkxs";
                DR["dedittime"] = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                DR["ideletetag"] = 0;
                DR["gid"] = irow.GetCell(35).StringCellValue;

                dt.Rows.Add(DR);
            }

            return dt;
        }


        protected DataTable PrepareMatrixDataTable(IWorkbook myworkbook, DataTable dt)
        {
            ISheet sheet = myworkbook.GetSheetAt(2);

            if (sheet.SheetName.Trim() != "SNX Summary_Standard Billing")
            {
                MessageBox.Show("Sheet Name or index have changed of Matrix");
                return null;
            }

            DataRow DR = null;

            IRow irow = null;

            int colNum = sheet.GetRow(0).LastCellNum + 1;

            for (int i = 1; i < sheet.LastRowNum + 1; i++)
            {
                DR = dt.NewRow();

                irow = sheet.GetRow(i);
                irow.GetCell(4).SetCellType(CellType.Numeric);
                if (irow.GetCell(0).CellType == CellType.Blank)
                {
                    return dt;
                }                
                DR["dchargemonth"] = ParaTime.Text.Trim();//getChargeMonthNow();
                DR["clocation"] = irow.GetCell(1).StringCellValue.Replace('\'', ' ');//this column may contain "\'"
                DR["mtotalcharge"] = irow.GetCell(4).NumericCellValue;
                DR["ceditby"] = "z003pkxs";
                DR["dedittime"] = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");

                dt.Rows.Add(DR);
            }

            return dt;
        }


        protected DataTable PrepareCIT_GIDDataTable(IWorkbook myworkbook, DataTable dt)
        {
            ISheet sheet = myworkbook.GetSheetAt(0);

            if (sheet.SheetName.Trim() != "Sheet1")
            {
                MessageBox.Show("Sheet Name or index have changed of CIT");
                return null;
            }

            int colNum = sheet.GetRow(0).LastCellNum + 1;
            if(colNum!=5&&colNum!=7){
                   MessageBox.Show("Columns of Excel have changed,please check and use column 4 or column 6 file"); 
                   return null; 
            }

            DataRow DR = null;

            IRow irow = null;

            //int colNum = sheet.GetRow(0).LastCellNum + 1;

            if (colNum == 7) {

                for (int i = 1; i < sheet.LastRowNum + 1; i++)
                {
                    DR = dt.NewRow();

                    irow = sheet.GetRow(i);
                    irow.GetCell(2).SetCellType(CellType.Numeric);
                    irow.GetCell(3).SetCellType(CellType.Numeric);
                               
                    DR["GID"] = irow.GetCell(1).StringCellValue;//getChargeMonthNow();
                    DR["Price"] = irow.GetCell(2).NumericCellValue;//this column may contain "\'"
                    DR["Ccomments"] = irow.GetCell(5).StringCellValue;
                    DR["quantity"] = irow.GetCell(3).NumericCellValue;
                    DR["servicename"] = irow.GetCell(0).StringCellValue;

                    dt.Rows.Add(DR);
                }
            }

            else if (colNum == 5)
            {

                for (int i = 1; i < sheet.LastRowNum + 1; i++)
                {
                    DR = dt.NewRow();

                    irow = sheet.GetRow(i);
                    irow.GetCell(2).SetCellType(CellType.Numeric);
                   
                    DR["GID"] = irow.GetCell(1).StringCellValue;
                    DR["Price"] = irow.GetCell(2).NumericCellValue;
                    DR["Ccomments"] = irow.GetCell(3).StringCellValue;
                    DR["quantity"] = 1;
                    DR["servicename"] = irow.GetCell(0).StringCellValue;

                    dt.Rows.Add(DR);
                }
            }

            
            return dt;
        }


        protected DataTable PrepareCDReportDataTable(IWorkbook myworkbook, DataTable dt, String FileNameWithoutSuffix)
        {
           

            ISheet sheet = myworkbook.GetSheetAt(0);

            String startdate = "", enddate = "";//, strperiod = "";// strlocation = "";
            

            if (sheet.SheetName.ToUpper().Trim() != "SHEET1")
            {
                MessageBox.Show("Sheet Name or index have changed of CDReport");
                return null;
            }

            int colNum = sheet.GetRow(0).LastCellNum + 1;
            if (colNum != 8)  //colNum != 5 && 
            {
                MessageBox.Show("Columns of Excel have changed,please check and use column 7 file");
                return null;
            }

            DataRow DR = null;

            IRow irow = null;

            //int colNum = sheet.GetRow(0).LastCellNum + 1;

            //DateTime dt;

            //dt = Convert.ToDateTime("2011/05/26", dtFormat);

            startdate = getFirstDayLastMonth(this.getChargeMonthNow(), dtFormat);
            enddate = getLastDayLastMonth(this.getChargeMonthNow(), dtFormat);


            if (colNum == 8)
            {

                for (int i = 1; i < sheet.LastRowNum + 1; i++)  // colnum 1,2,3 other regular info
                {
                    if (i == 12746)
                        i = 12746;

                    DR = dt.NewRow();

                    irow = sheet.GetRow(i);

                    //if (irow == null)
                    //    continue;


                    if(JudgeBlankRow(irow, 3))
                        continue;

                    irow.GetCell(5).SetCellType(CellType.Numeric);
                    irow.GetCell(6).SetCellType(CellType.Numeric);

                    CheckCellBlank(irow, 0, CellType.Blank);
                    CheckCellBlank(irow,1,CellType.Blank);
                    CheckCellBlank(irow, 2, CellType.Blank);
                    CheckCellBlank(irow, 3, CellType.Blank);
                    //irow.GetCell(3).SetCellType(CellType.Numeric);

                    //DR["GID"] = irow.GetCell(1).StringCellValue;//getChargeMonthNow();
                    //DR["Price"] = irow.GetCell(2).NumericCellValue;//this column may contain "\'"
                    //DR["Ccomments"] = irow.GetCell(5).StringCellValue;
                    //DR["quantity"] = irow.GetCell(3).NumericCellValue;
                    //DR["servicename"] = irow.GetCell(0).StringCellValue;


                    DR["csbs_id"] = irow.GetCell(0).StringCellValue;
                    DR["cperno"] = irow.GetCell(1).StringCellValue;
                    DR["cname"] = irow.GetCell(2).StringCellValue;
                    DR["cbuname"] = irow.GetCell(3).StringCellValue;
                    DR["ccostcenter"] = irow.GetCell(4).StringCellValue;
                    DR["mtotalprice"] = irow.GetCell(5).NumericCellValue;
                    DR["dstartdate"] = Convert.ToDateTime(startdate,dtFormat);
                    DR["denddate"] = Convert.ToDateTime(enddate, dtFormat);
                    DR["ccomments"] = FileNameWithoutSuffix;
                    DR["dchargemonth"] = getChargeMonthNow();
                    DR["ceditby"] = "CD_Report";
                    DR["dedittime"] = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                    DR["cremark"] = "S201"; //"S677";  //temp cserviceid
                    DR["itag"] = 0;

                    dt.Rows.Add(DR); 
                    /*                  
                     other part logic
                     */
                    DataRow NewDR = dt.NewRow();

                    NewDR["csbs_id"] = irow.GetCell(0).StringCellValue;
                    NewDR["cperno"] = irow.GetCell(1).StringCellValue;
                    NewDR["cname"] = irow.GetCell(2).StringCellValue;
                    NewDR["cbuname"] = irow.GetCell(3).StringCellValue;
                    NewDR["ccostcenter"] = irow.GetCell(4).StringCellValue;
                    NewDR["mtotalprice"] = irow.GetCell(6).NumericCellValue;
                    NewDR["dstartdate"] = Convert.ToDateTime(startdate, dtFormat);
                    NewDR["denddate"] = Convert.ToDateTime(enddate, dtFormat);
                    NewDR["ccomments"] = FileNameWithoutSuffix;
                    NewDR["dchargemonth"] = getChargeMonthNow();
                    NewDR["ceditby"] = "CD_Report";
                    NewDR["dedittime"] = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                    NewDR["cremark"] =  "S677";
                    NewDR["itag"] = 0;
      
                    dt.Rows.Add(NewDR);
                }
            }

            return dt;
        }


        /// <summary>
        /// travel pre LinkCell cells,if all blank,judge to blank line
        /// </summary>
        /// <param name="irow"> row obj need to judge</param>
        /// <param name="LinkCell"> cell number to judge blank line </param>
        /// <returns></returns>
        public bool JudgeBlankRow(IRow irow,int LinkCell) {

            int judgenumber = 0,blanksum=0;

            //Boolean result = false;

            if (irow == null)
                return true;

            if (LinkCell < 1)
                throw new Exception("Invalid LinkCell parameter");

            judgenumber = irow.LastCellNum;

            LinkCell = (LinkCell > judgenumber) ? judgenumber : LinkCell;

            for (int i = 0; i < LinkCell; i++)
            {
                if (irow.Cells[i].StringCellValue.Length==0)
                    blanksum++;
            }

            return (LinkCell==blanksum);
        }



        /// <summary>
        /// Dateformat : 'yyyy-mm-dd'
        /// </summary>
        /// <param name="Date"></param>
        /// <returns></returns>
        public static String getFirstDayLastMonth(String Date, DateTimeFormatInfo dtFormat)
        {
            DateTime datetime = Convert.ToDateTime(Date, dtFormat);
            return datetime.AddDays(1 - datetime.Day).AddMonths(-1).ToString();
        }
        /// <summary>
        /// Dateformat : 'yyyy-mm-dd'
        /// </summary>
        /// <param name="Date"></param>
        /// <returns></returns>
        public static String getLastDayLastMonth(String Date, DateTimeFormatInfo dtFormat)
        {
            DateTime datetime = Convert.ToDateTime(Date, dtFormat);
            return datetime.AddDays(1 - datetime.Day).AddDays(-1).ToString();
        }


        /// <summary>
        /// return null when column not contain "CHECKEDCC"
        /// </summary>
        /// <param name="ColArr"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        protected DataTable PrepareDataTableFromColumnName(String[] ColArr, DataTable dt)
        {
            Boolean CheckCC=false;

            dt.Columns.Add("id", Type.GetType("System.Int32"));

            //for (int i=0;i<ColArr.Length ;i++ ) {
            
            //}

            foreach (String colnum in ColArr) {
                dt.Columns.Add(colnum, Type.GetType("System.String"));

                if (colnum.Equals(InvalidCCColNumName))
                    CheckCC = true;
            }


            if (!CheckCC) {
                //dt.Clear();

                return null;
            }



                return dt;
        }


        protected DataTable PrepareCIT_EmailDataTable(IWorkbook myworkbook, DataTable dt)
        {
            ISheet sheet = myworkbook.GetSheetAt(0);

            if (sheet.SheetName.Trim() != "Sheet1")
            {
                MessageBox.Show("Sheet Name or index have changed of CIT Email");
                return null;
            }

            int colNum = sheet.GetRow(0).LastCellNum + 1;
            if (colNum != 5)
            {
                MessageBox.Show("Columns of Excel have changed,please check and use column 4 ");
                return null;
            }

            DataRow DR = null;

            IRow irow = null;

            //int colNum = sheet.GetRow(0).LastCellNum + 1;

            if (colNum == 5)
            {

                for (int i = 1; i < sheet.LastRowNum + 1; i++)
                {
                    DR = dt.NewRow();

                    irow = sheet.GetRow(i);
                    irow.GetCell(2).SetCellType(CellType.Numeric);
                    irow.GetCell(3).SetCellType(CellType.String);

                    DR["email"] = irow.GetCell(1).StringCellValue;
                    DR["Price"] = irow.GetCell(2).NumericCellValue;
                    DR["comments"] = irow.GetCell(3).StringCellValue;
                    //DR["quantity"] = 1;
                    DR["servicename"] = irow.GetCell(0).StringCellValue;

                    dt.Rows.Add(DR);
                }
            }


            return dt;
        }

        protected DataTable PrepareCIT_FlenderDataTable(IWorkbook myworkbook, DataTable dt)
        {
            ISheet sheet = myworkbook.GetSheetAt(0);

            if (sheet.SheetName.Trim() != "Sheet1")
            {
                MessageBox.Show("Sheet Name or index have changed of CIT Flender");
                return null;
            }

            int colNum = sheet.GetRow(0).LastCellNum;// +1;
            if (colNum != 15)
            {
                MessageBox.Show("Columns of Excel have changed,please check and use column 15 ");
                return null;
            }

            DataRow DR = null;

            IRow irow = null;

            //int colNum = sheet.GetRow(0).LastCellNum + 1;

            if (colNum == 15)
            {

                for (int i = 1; i < sheet.LastRowNum + 1; i++)
                {
                    DR = dt.NewRow();

                    irow = sheet.GetRow(i);
                    irow.GetCell(7).SetCellType(CellType.Numeric);
                    irow.GetCell(9).SetCellType(CellType.Numeric);
                    irow.GetCell(3).SetCellType(CellType.String);
                    irow.GetCell(1).SetCellType(CellType.String);
                    irow.GetCell(0).SetCellType(CellType.String);

                    DR["dchargemonth"] = ParaTime.Text.Trim();
                    DR["ccostcenter"] = irow.GetCell(0).StringCellValue;
                    DR["cperno"] = irow.GetCell(1).StringCellValue;
                    DR["cname"] = irow.GetCell(2).StringCellValue;
                    DR["csbs_id"] = irow.GetCell(3).StringCellValue;
                    DR["cfs_id"] = irow.GetCell(4).StringCellValue;
                    DR["cservicename"] = irow.GetCell(5).StringCellValue;
                    DR["ccomments"] = irow.GetCell(6).StringCellValue;
                    DR["munitprice"] = irow.GetCell(7).NumericCellValue;
                    DR["fquantity"] = irow.GetCell(8).NumericCellValue;
                    DR["mtotalprice"] = irow.GetCell(9).NumericCellValue;
                    DR["ceditby"] = irow.GetCell(10).StringCellValue;
                    DR["cservicetype"] = irow.GetCell(11).StringCellValue;
                    DR["ccioapptype"] = irow.GetCell(12).StringCellValue;
                    DR["creference"] = "";//irow.GetCell(13).StringCellValue;
                    DR["cversion"] = ""; //irow.GetCell(14).StringCellValue; always blank in excel
                    DR["cremark"] = "";//this colnum set null original
                    DR["ideletetag"] = 0;
                    DR["dedittime"] = DateTime.Now;

                    //DR["email"] = irow.GetCell(1).StringCellValue;
                    //DR["Price"] = irow.GetCell(2).NumericCellValue;
                    //DR["comments"] = irow.GetCell(3).StringCellValue;
                    ////DR["quantity"] = 1;
                    //DR["servicename"] = irow.GetCell(0).StringCellValue;

                    dt.Rows.Add(DR);
                }
            }


            return dt;
        }


        protected DataTable PrepareHrDataTable(IWorkbook myworkbook, DataTable dt)
        {
            ISheet sheet = myworkbook.GetSheetAt(0);

            DataRow DR = null;

            IRow irow = null;

            int colNum = sheet.GetRow(0).LastCellNum + 1;

            for (int i = 1; i < sheet.LastRowNum + 1; i++)
            {
                DR = dt.NewRow();

                irow = sheet.GetRow(i);
                //irow.GetCell(35).SetCellType(CellType.String);
               
                //irow.GetCell(1).SetCellType(CellType.Numeric);
                //if (irow.GetCell(0).CellType == CellType.Blank)

                //remove blank telephone  chinese char
                //if ((irow.GetCell(4).StringCellValue.Trim() == "") || checkChineseStr(irow.GetCell(4).StringCellValue.Trim()))
                //continue;
                // DR["id"] = i;

                DR["PersNo"] = irow.GetCell(0).StringCellValue;
                DR["CostCtr"] = irow.GetCell(8).StringCellValue;
                DR["BU"] = irow.GetCell(12).StringCellValue;
                DR["GID"] = irow.GetCell(15).StringCellValue;

                dt.Rows.Add(DR);
            }

            DataRow[] drArr=dt.Select("BU in ('SSMR','SSME','SSCX','SPA','SFAE')");

            DataTable dtNew = dt.Clone();
            for (int i = 0; i < drArr.Length; i++)
            {
                dtNew.ImportRow(drArr[i]);

            }


            //foreach (DataRow dr in drArr) {
            //    //dt.ImportRow(dr); 
            //    dt.Rows.Add(dr);
            //}

            return dtNew;
        }


        protected DataTable Readexcel(string path, string sheetname, Boolean HaveColumnName)
        {

            IWorkbook myworkbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;
            List<ICell> Cell_List = null;
            int count = 0;
            FileStream fs;
            //store column name
            string[] colname = null;

            int ColCount = 0;

            DataTable DT = new DataTable();
            DataRow DR = null;

            if (!checkFileExist(path))
                return null;

            fs = new FileStream(path, FileMode.Open, FileAccess.Read);

            if (path.IndexOf(".xlsx") > 0) // 2007版本
                myworkbook = new XSSFWorkbook(fs);
            else if (path.IndexOf(".xls") > 0) // 2003版本
                myworkbook = new HSSFWorkbook(fs);

            sheet = sheetname == null ? myworkbook.GetSheetAt(0) : myworkbook.GetSheet(sheetname);

            ColCount = sheet.GetRow(0).LastCellNum;
            colname = new string[ColCount];

            //get column value
            if (HaveColumnName)
            {
                row = sheet.GetRow(count);
                Cell_List = row.Cells;


                for (int i = 0; i < Cell_List.Count; i++)
                {
                    colname[i] = Cell_List[i].StringCellValue;
                }
                count++;
            }

            for (int i = count; i <= sheet.LastRowNum + 1; i++)
            {
                row = sheet.GetRow(count);
                if (row == null)
                    continue;

                DR = DT.NewRow();
                for (int j = row.FirstCellNum; j <= ColCount; j++)
                {
                    cell = row.GetCell(j);
                    //DR[j] = (row.GetCell(j) == null ? null : row.GetCell(j).StringCellValue.Replace('\'',' '));
                    if (cell == null)
                    {
                        DR[j] = null;
                    }
                    else
                    {
                        //读取Excel格式，根据格式读取数据类型
                        switch (cell.CellType)
                        {
                            case CellType.Blank: //空数据类型处理
                                DR[j] = "";
                                break;
                            case CellType.String: //字符串类型
                                DR[j] = cell.StringCellValue;
                                break;
                            case CellType.Numeric: //数字类型                                   
                                if (DateUtil.IsValidExcelDate(cell.NumericCellValue))
                                {
                                    DR[j] = cell.DateCellValue;
                                }
                                else
                                {
                                    DR[j] = cell.NumericCellValue;
                                }
                                break;
                        }
                    }

                    // replace '\'' to ' ' here
                }
                DT.Rows.Add(DR);
                //count++;            
            }
            return DT;
        }

        private Boolean checkFileExist(string path)
        {
            return File.Exists(path);
        }

        /// <summary>
        /// Check whether cell value is blank and set blank cell to ""
        /// </summary>
        /// <param name="irow"></param>
        /// <param name="index">real col index from 0</param>
        /// <param name="celltype">use to set blank cell type,always set to CellType.String</param>
        /// <returns></returns>
        protected Boolean CheckCellBlank(IRow irow,int index,CellType celltype) {
            try
            {
                if (irow.GetCell(index) == null)
                {
                    irow.CreateCell(index, celltype);
                    irow.GetCell(index).SetCellValue("");
                }
            }
            catch (Exception)
            {

                return false;
            }
            return true;
        }


        private void directline_Click(object sender, EventArgs e)
        {
            Boolean result = TravelDirectLineExcelInFloder(DefaultExcelPath + ParaTime.Text + directline_dir); 

            //MessageBoxButtons messButton = MessageBoxButtons.OK;

            if (result)
            {
                MessageBox.Show("import successfully");
                //MessageBox.Show("continue to excute 2 store procedures in DB!!!!");
                MessageBox.Show("continue to generate 2 excels from DB!!!!");

                //first directline file
                _ExcelName = "direct line " + ParaTime.Text.Trim() + "(result)";
                IWorkbook Myworkbook = null;
                DataTable dt = ExcuteQuerySP("sp_ba_n_step_12_alldirectline", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                CommonExecute(DefaultExcelPath + ParaTime.Text + directline_dir_out, _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);


                //second directline file
                _ExcelName = "direct line invalid data (status 1)" + ParaTime.Text.Trim();
                Myworkbook = null;
                dt = ExcuteQuerySP("sp_ba_n_step_12_invaliddirectline", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                CommonExecute(DefaultExcelPath + ParaTime.Text + directline_dir_out, _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);

                MessageBox.Show("Generate 2 Excel Files SuccessFully!!!!");
            }
            else
                MessageBox.Show("unknown error occour");      

        }


        private void Snx_Click(object sender, EventArgs e)
        {
            Boolean result = TravelSnxExcelInFloder(DefaultExcelPath + ParaTime.Text + snx_dir);

            //MessageBoxButtons messButton = MessageBoxButtons.OK;

            if (result)
            {
                MessageBox.Show("import Snx and Bcp successfully");


                MessageBox.Show("continue to generate Invalid cc File of SNX!!!!");

                //first directline file
                //_ExcelName = "direct line " + ParaTime.Text.Trim() + "(result)";
                //IWorkbook Myworkbook = null;
                //DataTable dt = ExcuteQuerySP("sp_ba_n_step_12_alldirectline", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                //CommonExecute(DefaultExcelPath, _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);


                ////second directline file
                //_ExcelName = "direct line invalid data (status 1)" + ParaTime.Text.Trim();
                //Myworkbook = null;
                //dt = ExcuteQuerySP("sp_ba_n_step_12_invaliddirectline", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                //CommonExecute(DefaultExcelPath, _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);

            }
            else
                MessageBox.Show("unknown error occour");    
        }


        protected Boolean DataTable2DB(SqlBulkCopy bulkcopy, DataTable dt)
        {
            //SqlBulkCopy bulkcopy = new SqlBulkCopy(DBconStr);

            try
            {
                bulkcopy.WriteToServer(dt);
            }
            catch (Exception ee)
            {
                //throw new Exception ee;
               //return false;
            }

            return true;
        }

        protected SqlBulkCopy initDirectLine_SqlBulkCopy(SqlBulkCopy bulkcopy, String DBconStr, String TableName)
        {
            bulkcopy = new SqlBulkCopy(DBconStr);
            bulkcopy.DestinationTableName = TableName;
            bulkcopy.BatchSize = 750;

            bulkcopy.ColumnMappings.Add("id", "id");
            bulkcopy.ColumnMappings.Add("dchargemonth", "dchargemonth");
            bulkcopy.ColumnMappings.Add("telephone", "cfullnumber");
            bulkcopy.ColumnMappings.Add("starttime", "ddatefrom");
            bulkcopy.ColumnMappings.Add("stoptime", "ddateto");
            bulkcopy.ColumnMappings.Add("totalprice", "mtotalprice");
            bulkcopy.ColumnMappings.Add("istatus", "istatus");
            bulkcopy.ColumnMappings.Add("ceditby", "ceditby");
            bulkcopy.ColumnMappings.Add("dedittime", "dedittime");
            return bulkcopy;
        }

        protected SqlBulkCopy initSnx_SqlBulkCopy(SqlBulkCopy bulkcopy, String DBconStr, String TableName)
        {
            bulkcopy = new SqlBulkCopy(DBconStr);
            bulkcopy.DestinationTableName = TableName;
            bulkcopy.BatchSize = 750;

            bulkcopy.ColumnMappings.Add("id", "id");
            bulkcopy.ColumnMappings.Add("dchargemonth", "dchargemonth");
            bulkcopy.ColumnMappings.Add("ccostcenter", "ccostcenter");
            bulkcopy.ColumnMappings.Add("cperno", "cperno");
            bulkcopy.ColumnMappings.Add("cname", "cname");
            bulkcopy.ColumnMappings.Add("clocation", "clocation");
            bulkcopy.ColumnMappings.Add("corganization", "corganization");
            bulkcopy.ColumnMappings.Add("ceditby", "ceditby");
            bulkcopy.ColumnMappings.Add("dedittime", "dedittime");
            bulkcopy.ColumnMappings.Add("GID", "GID");
            return bulkcopy;
        }


        protected SqlBulkCopy initBcp_SqlBulkCopy(SqlBulkCopy bulkcopy, String DBconStr, String TableName)
        {
            bulkcopy = new SqlBulkCopy(DBconStr);
            bulkcopy.DestinationTableName = TableName;
            bulkcopy.BatchSize = 750;

            bulkcopy.ColumnMappings.Add("id", "id");
            bulkcopy.ColumnMappings.Add("dchargemonth", "dchargemonth");          
            bulkcopy.ColumnMappings.Add("cperno", "cperno");
            bulkcopy.ColumnMappings.Add("cname", "cname");
            bulkcopy.ColumnMappings.Add("cemail", "cemail");
            bulkcopy.ColumnMappings.Add("ccostcenter", "ccostcenter");
            bulkcopy.ColumnMappings.Add("ideletetag", "ideletetag");
            bulkcopy.ColumnMappings.Add("ceditby", "ceditby");
            bulkcopy.ColumnMappings.Add("dedittime", "dedittime");
            bulkcopy.ColumnMappings.Add("gid", "gid");
            return bulkcopy;
        }

        protected SqlBulkCopy initMatrix_SqlBulkCopy(SqlBulkCopy bulkcopy, String DBconStr, String TableName)
        {
            bulkcopy = new SqlBulkCopy(DBconStr);
            bulkcopy.DestinationTableName = TableName;
            bulkcopy.BatchSize = 750;

            bulkcopy.ColumnMappings.Add("id", "id");
            bulkcopy.ColumnMappings.Add("dchargemonth", "dchargemonth");
            bulkcopy.ColumnMappings.Add("clocation", "clocation");
            bulkcopy.ColumnMappings.Add("mtotalcharge", "mtotalcharge");
            bulkcopy.ColumnMappings.Add("ceditby", "ceditby");
            bulkcopy.ColumnMappings.Add("dedittime", "dedittime");
            return bulkcopy;
        }

        protected SqlBulkCopy initHr_SqlBulkCopy(SqlBulkCopy bulkcopy, String DBconStr, String TableName)
        {
            bulkcopy = new SqlBulkCopy(DBconStr);
            bulkcopy.DestinationTableName = TableName;
            bulkcopy.BatchSize = 750;
            bulkcopy.ColumnMappings.Add("PersNo", "PersNo");
            bulkcopy.ColumnMappings.Add("CostCtr", "CostCtr");
            bulkcopy.ColumnMappings.Add("BU", "BU");
            bulkcopy.ColumnMappings.Add("GID", "GID");
            return bulkcopy;
        }

        protected SqlBulkCopy initCIT_GID_SqlBulkCopy(SqlBulkCopy bulkcopy, String DBconStr, String TableName)
        {
            bulkcopy = new SqlBulkCopy(DBconStr);
            bulkcopy.DestinationTableName = TableName;
            bulkcopy.BatchSize = 750;
            bulkcopy.ColumnMappings.Add("servicename", "servicename");
            bulkcopy.ColumnMappings.Add("gid", "gid");
            bulkcopy.ColumnMappings.Add("quantity", "quantity");
            bulkcopy.ColumnMappings.Add("price", "price");
            bulkcopy.ColumnMappings.Add("ccomments", "ccomments");
            return bulkcopy;
        }

        protected SqlBulkCopy initCDReportSqlBulkCopy(SqlBulkCopy bulkcopy, String DBconStr, String TableName)
        {
            bulkcopy = new SqlBulkCopy(DBconStr);
            bulkcopy.DestinationTableName = TableName;
            bulkcopy.BatchSize = 750;
            //bulkcopy.ColumnMappings.Add("servicename", "servicename");
            //bulkcopy.ColumnMappings.Add("gid", "gid");
            //bulkcopy.ColumnMappings.Add("quantity", "quantity");
            //bulkcopy.ColumnMappings.Add("price", "price");
            //bulkcopy.ColumnMappings.Add("ccomments", "ccomments");

            bulkcopy.ColumnMappings.Add("id", "id");
            bulkcopy.ColumnMappings.Add("csbs_id", "csbs_id");
            bulkcopy.ColumnMappings.Add("cperno", "cperno");
            bulkcopy.ColumnMappings.Add("cname", "cname");
            bulkcopy.ColumnMappings.Add("cbuname", "cbuname");
            bulkcopy.ColumnMappings.Add("ccostcenter", "ccostcenter");
            bulkcopy.ColumnMappings.Add("mtotalprice", "mtotalprice");
            bulkcopy.ColumnMappings.Add("dstartdate", "dstartdate");
            bulkcopy.ColumnMappings.Add("denddate", "denddate");
            bulkcopy.ColumnMappings.Add("ccomments", "ccomments");
            bulkcopy.ColumnMappings.Add("dchargemonth", "dchargemonth");
            bulkcopy.ColumnMappings.Add("ceditby", "ceditby");
            bulkcopy.ColumnMappings.Add("dedittime", "dedittime");
            bulkcopy.ColumnMappings.Add("cremark", "cremark");
            bulkcopy.ColumnMappings.Add("itag", "itag");

            return bulkcopy;
        }



        protected SqlBulkCopy initCIT_Email_SqlBulkCopy(SqlBulkCopy bulkcopy, String DBconStr, String TableName)
        {
            bulkcopy = new SqlBulkCopy(DBconStr);
            bulkcopy.DestinationTableName = TableName;
            bulkcopy.BatchSize = 750;
            bulkcopy.ColumnMappings.Add("id", "id");
            bulkcopy.ColumnMappings.Add("servicename", "ServiceName");
            bulkcopy.ColumnMappings.Add("email", "email");
            bulkcopy.ColumnMappings.Add("price", "price");
            bulkcopy.ColumnMappings.Add("comments", "comments");
            return bulkcopy;
        }

        protected SqlBulkCopy initCIT_Flender_SqlBulkCopy(SqlBulkCopy bulkcopy, String DBconStr, String TableName)
        {
            bulkcopy = new SqlBulkCopy(DBconStr);
            bulkcopy.DestinationTableName = TableName;
            bulkcopy.BatchSize = 750;
            bulkcopy.ColumnMappings.Add("id", "id");
            bulkcopy.ColumnMappings.Add("dchargemonth", "dchargemonth");
            bulkcopy.ColumnMappings.Add("ccostcenter", "ccostcenter");
            bulkcopy.ColumnMappings.Add("cperno", "cperno");
            bulkcopy.ColumnMappings.Add("cname", "cname");
            bulkcopy.ColumnMappings.Add("csbs_id", "csbs_id");
            bulkcopy.ColumnMappings.Add("cfs_id", "cfs_id");
            bulkcopy.ColumnMappings.Add("cservicename", "cservicename");
            bulkcopy.ColumnMappings.Add("ccomments", "ccomments");
            bulkcopy.ColumnMappings.Add("cservicetype", "cservicetype");
            bulkcopy.ColumnMappings.Add("munitprice", "munitprice");
            bulkcopy.ColumnMappings.Add("fquantity", "fquantity");
            bulkcopy.ColumnMappings.Add("mtotalprice", "mtotalprice");
            bulkcopy.ColumnMappings.Add("creference", "creference");
            bulkcopy.ColumnMappings.Add("cversion", "cversion");
            bulkcopy.ColumnMappings.Add("ceditby", "ceditby");
            bulkcopy.ColumnMappings.Add("dedittime", "dedittime");
            bulkcopy.ColumnMappings.Add("cremark", "cremark");
            bulkcopy.ColumnMappings.Add("ideletetag", "ideletetag");
            bulkcopy.ColumnMappings.Add("ccioapptype", "ccioapptype");
            return bulkcopy;
        }

        protected DataTable initDirectLine_Datatable(DataTable dt)
        {

            dt.Columns.Add("id", Type.GetType("System.Int32"));
            dt.Columns.Add("dchargemonth", Type.GetType("System.DateTime"));
            dt.Columns.Add("telephone", Type.GetType("System.String"));
            dt.Columns.Add("starttime", Type.GetType("System.DateTime"));
            dt.Columns.Add("stoptime", Type.GetType("System.DateTime"));
            dt.Columns.Add("totalprice", Type.GetType("System.Double"));
            dt.Columns.Add("istatus", Type.GetType("System.Double"));
            dt.Columns.Add("ceditby", Type.GetType("System.String"));
            dt.Columns.Add("dedittime", Type.GetType("System.DateTime"));

            return dt;
        }


        protected DataTable initSNX_Datatable(DataTable dt)
        {

            dt.Columns.Add("id", Type.GetType("System.Int32"));
            dt.Columns.Add("dchargemonth", Type.GetType("System.String"));
            dt.Columns.Add("ccostcenter", Type.GetType("System.String"));
            dt.Columns.Add("cperno", Type.GetType("System.String"));
            dt.Columns.Add("cname", Type.GetType("System.String"));
            dt.Columns.Add("clocation", Type.GetType("System.String"));
            dt.Columns.Add("corganization", Type.GetType("System.String"));
            dt.Columns.Add("ceditby", Type.GetType("System.String"));
            dt.Columns.Add("dedittime", Type.GetType("System.DateTime"));
            dt.Columns.Add("GID", Type.GetType("System.String"));

            return dt;
        }

        protected DataTable initBcp_Datatable(DataTable dt)
        {

            dt.Columns.Add("id", Type.GetType("System.Int32"));
            dt.Columns.Add("dchargemonth", Type.GetType("System.String"));           
            dt.Columns.Add("cperno", Type.GetType("System.String"));
            dt.Columns.Add("cname", Type.GetType("System.String"));
            dt.Columns.Add("cemail", Type.GetType("System.String"));
            dt.Columns.Add("ccostcenter", Type.GetType("System.String"));
            dt.Columns.Add("ideletetag", Type.GetType("System.Int32"));
            dt.Columns.Add("ceditby", Type.GetType("System.String"));
            dt.Columns.Add("dedittime", Type.GetType("System.DateTime"));
            dt.Columns.Add("gid", Type.GetType("System.String"));

            return dt;
        }
        protected DataTable initMatrix_Datatable(DataTable dt)
        {
            dt.Columns.Add("id", Type.GetType("System.Int32"));
            dt.Columns.Add("dchargemonth", Type.GetType("System.String"));
            dt.Columns.Add("clocation", Type.GetType("System.String"));
            dt.Columns.Add("mtotalcharge", Type.GetType("System.Double"));
            dt.Columns.Add("ceditby", Type.GetType("System.String"));
            dt.Columns.Add("dedittime", Type.GetType("System.DateTime"));
           

            return dt;
        }

        protected DataTable initCIT_GID_Datatable(DataTable dt)
        {

            dt.Columns.Add("id", Type.GetType("System.Int32"));
            dt.Columns.Add("GID", Type.GetType("System.String"));
            dt.Columns.Add("Price", Type.GetType("System.Double"));
            dt.Columns.Add("Ccomments", Type.GetType("System.String"));
            dt.Columns.Add("quantity", Type.GetType("System.String"));
            dt.Columns.Add("servicename", Type.GetType("System.String"));

            return dt;
        }

        protected DataTable initCDReport_Datatable(DataTable dt)
        {

            dt.Columns.Add("id", Type.GetType("System.Int32"));
            dt.Columns.Add("csbs_id", Type.GetType("System.String"));
            dt.Columns.Add("cperno", Type.GetType("System.String"));
            dt.Columns.Add("cname", Type.GetType("System.String"));
            dt.Columns.Add("cbuname", Type.GetType("System.String"));
            dt.Columns.Add("ccostcenter", Type.GetType("System.String"));
            dt.Columns.Add("mtotalprice", Type.GetType("System.String"));
            dt.Columns.Add("dstartdate", Type.GetType("System.DateTime"));
            dt.Columns.Add("denddate", Type.GetType("System.DateTime"));
            dt.Columns.Add("ccomments", Type.GetType("System.String"));
            dt.Columns.Add("dchargemonth", Type.GetType("System.String"));
            dt.Columns.Add("ceditby", Type.GetType("System.String"));
            dt.Columns.Add("dedittime", Type.GetType("System.DateTime"));
            dt.Columns.Add("cremark", Type.GetType("System.String"));
            dt.Columns.Add("itag", Type.GetType("System.Int32"));

            return dt;
        }


        protected DataTable initCIT_Email_Datatable(DataTable dt)
        {

            dt.Columns.Add("id", Type.GetType("System.Int32"));
            dt.Columns.Add("servicename", Type.GetType("System.String"));
            dt.Columns.Add("email", Type.GetType("System.String"));
            dt.Columns.Add("price", Type.GetType("System.Double"));
            dt.Columns.Add("comments", Type.GetType("System.String"));

            return dt;
        }

         protected DataTable initCIT_Flender_Datatable(DataTable dt)
        {

            dt.Columns.Add("id", Type.GetType("System.Int32"));
            dt.Columns.Add("dchargemonth", Type.GetType("System.String"));
            dt.Columns.Add("ccostcenter", Type.GetType("System.String"));
            dt.Columns.Add("cperno", Type.GetType("System.String"));
            dt.Columns.Add("cname", Type.GetType("System.String"));
            dt.Columns.Add("csbs_id", Type.GetType("System.String"));
            dt.Columns.Add("cfs_id", Type.GetType("System.String"));
            dt.Columns.Add("cservicename", Type.GetType("System.String"));
            dt.Columns.Add("ccomments", Type.GetType("System.String"));
            dt.Columns.Add("cservicetype", Type.GetType("System.String"));
            dt.Columns.Add("munitprice", Type.GetType("System.Double"));
            dt.Columns.Add("fquantity", Type.GetType("System.Double"));
            dt.Columns.Add("mtotalprice", Type.GetType("System.Double"));
            dt.Columns.Add("creference", Type.GetType("System.String"));
            dt.Columns.Add("ccioapptype", Type.GetType("System.String"));
            dt.Columns.Add("cremark", Type.GetType("System.String"));
            dt.Columns.Add("cversion", Type.GetType("System.String"));
            dt.Columns.Add("ceditby", Type.GetType("System.String"));
            dt.Columns.Add("ideletetag", Type.GetType("System.Int32"));
            dt.Columns.Add("dedittime", Type.GetType("System.String"));
            

            return dt;
        }

        protected DataTable initHr_Datatable(DataTable dt)
        {
            dt.Columns.Add("id", Type.GetType("System.Int32"));
            dt.Columns.Add("PersNo", Type.GetType("System.String"));
            dt.Columns.Add("CostCtr", Type.GetType("System.String"));
            dt.Columns.Add("BU", Type.GetType("System.String"));
            dt.Columns.Add("GID", Type.GetType("System.String"));
            return dt;
        }

        protected Boolean checkChineseStr(String BechcekStr)
        {

            char[] c = BechcekStr.Trim().ToCharArray();
            for (int i = 0; i < c.Length; i++)
            {
                if (c[i] >= 0x4e00 && c[i] <= 0x9fbb)
                {
                    return true;
                }
            }
            return false;
        }


        private void IT_fault_data_Click(object sender, EventArgs e)
        {
            _SheetName_Arr = new String[] { "Q2CD_CALL_CLOSED", "Q2CD_Contract", "Q2CD_HW_CONFIG", "Q2CD_SERV_CONFIG", "Volume", "Volume_Unit_Price" };
            _ExcelName = "IT Fault Data_" + ParaTime.Text.Trim();

            IWorkbook Myworkbook = null;
            Boolean isEnd = false;
            string sheetname = "";
            for (int i = 0; i < _SheetName_Arr.Length; i++)
            {
                sheetname = _SheetName_Arr[i];
                if (i == _SheetName_Arr.Length - 1)
                    isEnd = true;

                DataTable dt = ExcuteQuerySP(GetRegularReportSpName("sp_ba_n_step_14_", sheetname), "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                dt.TableName = sheetname;
                CommonExecute(DefaultExcelPath + ParaTime.Text + "\\report", _ExcelName, sheetname, dt, Myworkbook, msg, isEnd, out Myworkbook);
            }
        }

        private void newcc_Click(object sender, EventArgs e)
        {
            _ExcelName = "NewCC_" + ParaTime.Text.Trim();
            IWorkbook Myworkbook = null;
            DataTable dt = null;
            BackgroundWorker mbg =new BackgroundWorker();
            mbg.DoWork+=new DoWorkEventHandler(m_bgWorker_DoWork);
            mbg.RunWorkerCompleted += new RunWorkerCompletedEventHandler(m_bgWorker_RunWorkerCompleted);
            //1 = a;    
            mbg.RunWorkerAsync();
            newcc.Enabled = false;
            while (mbg.IsBusy)
            {
                MainProgressBar.Increment(1);
                // Keep UI messages moving, so the form remains 
                // responsive during the asynchronous operation.
                Application.DoEvents();
            }
            dt = _dt;

                //dt= ExcuteQuerySP("sp_ba_n_step_05_newcc", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
            CommonExecute(DefaultExcelPath + ParaTime.Text+"\\report", _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);
        }

        private void SnxTempInvalidCC_Click(object sender, EventArgs e)
        {
            _ExcelName = "SNX invalid cc-" + ParaTime.Text.Trim();
            IWorkbook Myworkbook = null;
            DataTable dt = ExcuteQuerySP("sp_ba_n_step_18_Snx_cleanTempTableCC", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
            CommonExecute(DefaultExcelPath + ParaTime.Text + snx_dir_out, _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);
        }

        private void SnxInvalidCC_Click(object sender, EventArgs e)
        {
            _ExcelName = "SNX invalid cc-" + ParaTime.Text.Trim()+"_DoubleCheck";
            IWorkbook Myworkbook = null;
            DataTable dt = ExcuteQuerySP("sp_ba_n_step_18_Snx_cleanChargeTableCC", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
            CommonExecute(DefaultExcelPath + ParaTime.Text + snx_dir_out, _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);
        }


        private void ImportMatrix_Click(object sender, EventArgs e)
        {
            Boolean result = TravelMatrixExcelInFloder(DefaultExcelPath + ParaTime.Text + matrix_dir);

            //MessageBoxButtons messButton = MessageBoxButtons.OK;

            if (result)
            {
                MoveImportedFile(DefaultExcelPath + ParaTime.Text + matrix_dir);
                MessageBox.Show("import Matrix successfully");

                //MessageBox.Show("continue to generate Invalid cc File of SNX!!!!");

                //first directline file
                //_ExcelName = "direct line " + ParaTime.Text.Trim() + "(result)";
                //IWorkbook Myworkbook = null;
                //DataTable dt = ExcuteQuerySP("sp_ba_n_step_12_alldirectline", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                //CommonExecute(DefaultExcelPath, _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);


                ////second directline file
                //_ExcelName = "direct line invalid data (status 1)" + ParaTime.Text.Trim();
                //Myworkbook = null;
                //dt = ExcuteQuerySP("sp_ba_n_step_12_invaliddirectline", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                //CommonExecute(DefaultExcelPath, _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);

            }
            else
                MessageBox.Show("unknown error occour"); 
        }

        protected bool Initfloder() {
            bool result = true;


            result = (result && checkFolderExist(DefaultExcelPath + ParaTime.Text + snx_dir));

            result = (result && checkFolderExist(DefaultExcelPath + ParaTime.Text + directline_dir));

            result = (result && checkFolderExist(DefaultExcelPath + ParaTime.Text + matrix_dir));

            result = (result && checkFolderExist(DefaultExcelPath + ParaTime.Text + snx_dir_out));

            result = (result && checkFolderExist(DefaultExcelPath + ParaTime.Text + directline_dir_out));

            result = (result && checkFolderExist(DefaultExcelPath + ParaTime.Text + "\\report"));

            result = (result && checkFolderExist(DefaultExcelPath + ParaTime.Text + cit_email_dir));

            result = (result && checkFolderExist(DefaultExcelPath + ParaTime.Text + cit_gid_dir));

            result = (result && checkFolderExist(DefaultExcelPath + ParaTime.Text + cit_hr_dir));

            result = (result && checkFolderExist(DefaultExcelPath + ParaTime.Text + cit_hr_dir_out));

            result = (result && checkFolderExist(DefaultExcelPath + ParaTime.Text + cit_flender_dir));

            result = (result && checkFolderExist(DefaultExcelPath + ParaTime.Text + Imported_moved));

            result = (result && checkFolderExist(DefaultExcelPath + ParaTime.Text + check_CC));

            result = (result && checkFolderExist(DefaultExcelPath + ParaTime.Text + CDReport_dir));

            result = (result && checkFolderExist(DefaultExcelPath + ParaTime.Text + check_BU));
            
            
            return result;
        }

        protected Hashtable InitFileHashtable(Hashtable FileTable)
        {
            if (FileTable==null)
                FileTable=new Hashtable();
            if(FileTable.ContainsKey(ImportedFileKey))
                FileTable.Remove(ImportedFileKey);
            FileTable.Add(ImportedFileKey,DefaultExcelPath + ParaTime.Text + Imported_moved);

            return FileTable;
        
        }

        private void ParaTime_TextChanged(object sender, EventArgs e)
        {
            //if (ParaTime.Text.Trim().Length == 10)
            //{
            //    Initfloder();
            //    FileTable = InitFileHashtable(this.FileTable);
            //}
                
        }

        private void CIT_with_GID_Click(object sender, EventArgs e)
        {
            Boolean result = TravelCIT_GIDExcelInFloder(DefaultExcelPath + ParaTime.Text + cit_gid_dir);

            //MessageBoxButtons messButton = MessageBoxButtons.OK;

            if (result)
            {
                MessageBox.Show("import CIT_GID File successfully");

                //MessageBox.Show("continue to generate Invalid cc File of SNX!!!!");

                //first directline file
                //_ExcelName = "direct line " + ParaTime.Text.Trim() + "(result)";
                //IWorkbook Myworkbook = null;
                //DataTable dt = ExcuteQuerySP("sp_ba_n_step_12_alldirectline", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                //CommonExecute(DefaultExcelPath, _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);


                ////second directline file
                //_ExcelName = "direct line invalid data (status 1)" + ParaTime.Text.Trim();
                //Myworkbook = null;
                //dt = ExcuteQuerySP("sp_ba_n_step_12_invaliddirectline", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                //CommonExecute(DefaultExcelPath, _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);

            }
            else
                MessageBox.Show("unknown error occour"); 

            //ImportCIT_GIDExcel
        }

        private void CIT_with_EMAIL_Click(object sender, EventArgs e)
        {
            Boolean result = TravelCIT_EmailExcelInFloder(DefaultExcelPath + ParaTime.Text + cit_email_dir);

            //MessageBoxButtons messButton = MessageBoxButtons.OK;

            if (result)
            {
                MessageBox.Show("import CIT_Email File successfully");

                //MessageBox.Show("continue to generate Invalid cc File of SNX!!!!");

                //first directline file
                //_ExcelName = "direct line " + ParaTime.Text.Trim() + "(result)";
                //IWorkbook Myworkbook = null;
                //DataTable dt = ExcuteQuerySP("sp_ba_n_step_12_alldirectline", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                //CommonExecute(DefaultExcelPath, _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);


                ////second directline file
                //_ExcelName = "direct line invalid data (status 1)" + ParaTime.Text.Trim();
                //Myworkbook = null;
                //dt = ExcuteQuerySP("sp_ba_n_step_12_invaliddirectline", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                //CommonExecute(DefaultExcelPath, _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);

            }
            else
                MessageBox.Show("unknown error occour"); 
        }

        private void HR_Click(object sender, EventArgs e)
        {
            String[] File=Directory.GetFiles(DefaultExcelPath+ParaTime.Text.Trim()+cit_hr_dir);

            
            foreach (string f in File)
            {
                if (ImportHrExcel(f))
                {
                    MoveImportedFile(f);
                    MessageBox.Show("HR Successfully in TempHR");
                }

                else {
                    MessageBox.Show("HR Failed!");
                }
            }
                
        }

        private void generate_hr_txt_Click(object sender, EventArgs e)
        {
             int  hrflag = 0;
             DataTable dt=null;
             dt = ExcuteQuerySPOutFlag("sp_ba_n_step_04_GetHrInfo", "", null, _configBA, "@flag", out  hrflag);
             String FullFileName = DefaultExcelPath + ParaTime.Text.Trim() + cit_hr_dir_out+ParaTime.Text.Trim()+".txt"; 
            //true
             if (hrflag == 1) {

                 try
                 {
                     if (!File.Exists(FullFileName))
                     {
                         File.Create(FullFileName).Dispose();
                     }
                 }
                 catch (Exception ee)
                 {                   
                    MessageBox.Show("Create txt issue! "+ee.ToString());
                 }

                 FileStream fs = new FileStream(FullFileName,FileMode.OpenOrCreate);
                 StreamWriter sw = new StreamWriter(fs);

                 sw.WriteLine("@echo off");
                 sw.WriteLine(" ");
                 foreach(DataRow dr in dt.Rows){
                     sw.WriteLine(BatchQ_config + "\"update(f_chq_call_entry,customer," + dr["qpkey"] + ",customer.costlocation=" + dr["CostCtr"] + "&customer.customerid=" + dr["PersNo"] + ")\""); //1 = a;
                 }
                 sw.WriteLine(" ");
                 sw.WriteLine("pause");

                 sw.Close();
                 sw.Dispose();

                 MessageBox.Show("Generate Hr File Successfully");
             }
               
             else if(hrflag == 0){

                 MessageBox.Show("Error Happened");
             }
        }
        //1=a;
        private void Import_CIT_flender_Click(object sender, EventArgs e)
        {
            Boolean result = TravelCIT_FlenderExcelInFloder(DefaultExcelPath + ParaTime.Text + cit_flender_dir);

            //MessageBoxButtons messButton = MessageBoxButtons.OK;

            if (result)
            {
                MessageBox.Show("import CIT_flender File successfully");

                //MessageBox.Show("continue to generate Invalid cc File of SNX!!!!");

                //first directline file
                //_ExcelName = "direct line " + ParaTime.Text.Trim() + "(result)";
                //IWorkbook Myworkbook = null;
                //DataTable dt = ExcuteQuerySP("sp_ba_n_step_12_alldirectline", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                //CommonExecute(DefaultExcelPath, _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);


                ////second directline file
                //_ExcelName = "direct line invalid data (status 1)" + ParaTime.Text.Trim();
                //Myworkbook = null;
                //dt = ExcuteQuerySP("sp_ba_n_step_12_invaliddirectline", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                //CommonExecute(DefaultExcelPath, _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);

            }
            else
                MessageBox.Show("unknown error occour"); 
        }
        /// <summary>
        /// aaa.xls  bbb.txt
        /// </summary>
        /// <param name="FullFileName"></param>
        /// <returns> FileName as hashtable key like test.txt</returns>
        protected String GetFileNameFromFullFileName(String FullFileName){
            string[] files=FullFileName.Split('\\');
            return files[files.Length-1];
        }


        /// <summary>
        /// aaa   bbb
        /// </summary>
        /// <param name="FullFileName"></param>
        /// <returns> FileName as hashtable key like test.txt</returns>
        protected String GetFileNameFromFullFileNameWithoutSuffix(String FullFileName)
        {
            string[] files = FullFileName.Split('\\');
            files = files[files.Length - 1].Split('.');

            return files[0];
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="FullFileName"></param>
        /// <returns> FileName as hashtable key like test.txt</returns>
        protected String ChangeFilePath(String FullFileName,String ChangePart)
        {
            string[] files = FullFileName.Split('\\');

            string path = "";

            files[files.Length - 1] = ChangePart+files[files.Length - 1];

            for (int i = 0; i < files.Length;i++ )
            {
                path = path + files[i];

                if (i != files.Length-1)
                    path = path + '\\';
            }

            return path;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Filetable">filename</param>
        /// <param name="FullFileName">file folder path</param>
        /// <returns></returns>
        protected Hashtable SetFullFileNameIntoHashtable(Hashtable Filetable, String FullFileName)
        {
            string[] files = FullFileName.Split('\\');
            //key
            String FileName = null;
            //value;
            String FolderPath = null;
            FileName = files[files.Length - 1];

            for (int i=0;i<files.Length-1 ;i++ ) {
                FolderPath = FolderPath + files[i]+"\\";
            }

            Filetable.Add(FileName, FolderPath);
                       
            //return files[files.Length - 1];
            return Filetable;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="CopyFullFileName">Full File Name</param>
        /// <returns></returns>
        protected Boolean MoveImportedFile(String CopyFullFileName) { 
            
            if(FileTable==null){
                throw new Exception("Uncatch init HashTable Failed Exception");
                //return false;
            }
            String FileName=GetFileNameFromFullFileName(CopyFullFileName);

            String ImportedFileFolder = (String)FileTable[ImportedFileKey];

            String SourceFileFolder=(String)FileTable[FileName];

            try
            {
                File.Move(SourceFileFolder + FileName, ImportedFileFolder + FileName);
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        private void LoaclDB_CheckedChanged(object sender, EventArgs e)
        {

            if (LoaclDB.Checked) 
                _configBA = _configTest;
        }

        private void CheckExcelCC_Click(object sender, EventArgs e)
        {
            TravelCheckCCExcelInFloder(DefaultExcelPath + ParaTime.Text + check_CC);
            

            /*
              find colnum named CHECKCC
             
             */

        }

        private void ChangeInitConfig_Click(object sender, EventArgs e)
        {
            bool MonthConfig = false, FolderConfig=false,result=false;
            MonthConfig=MonthConfigChange();
            FolderConfig = FolderConfigChange();
            result=MonthConfig&&FolderConfig;

            if (result)
                MessageBox.Show("Change Month and folder successfully!");

            else if (!result)
                MessageBox.Show("Change Month and folder failed!Please Check your input!");

        }


        public bool MonthConfigChange(){//(Winformobj wobj) {
            bool result = false;
            DefaultExcelPath = ExcelStorePath.Text.Trim();
            result = Initfloder();
            FileTable = InitFileHashtable(this.FileTable);
            return result; 
        }


        public bool FolderConfigChange()
        {
            bool result = false;
            if (ParaTime.Text.Trim().Length == 10)
            {
                result=Initfloder();
                FileTable = InitFileHashtable(this.FileTable);
            }

            return result;
        }

        private void CD_Report_New_Click(object sender, EventArgs e)
        {
            Boolean result = TravelCDReportExcelInFloder(DefaultExcelPath + ParaTime.Text + CDReport_dir);

            if (result)
            {
                MessageBox.Show("import CDReport File successfully");

                //MessageBox.Show("continue to generate Invalid cc File of SNX!!!!");

                //first directline file
                //_ExcelName = "direct line " + ParaTime.Text.Trim() + "(result)";
                //IWorkbook Myworkbook = null;
                //DataTable dt = ExcuteQuerySP("sp_ba_n_step_12_alldirectline", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                //CommonExecute(DefaultExcelPath, _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);


                ////second directline file
                //_ExcelName = "direct line invalid data (status 1)" + ParaTime.Text.Trim();
                //Myworkbook = null;
                //dt = ExcuteQuerySP("sp_ba_n_step_12_invaliddirectline", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
                //CommonExecute(DefaultExcelPath, _ExcelName, "Sheet1", dt, Myworkbook, msg, true, out Myworkbook);

            }
            else
                MessageBox.Show("unknown error occour");
        }

        private void m_bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            e.Result = ExcuteQuerySP("sp_ba_n_step_05_newcc", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
        }

        private void m_bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            _dt=(DataTable)e.Result; //= ExcuteQuerySP("sp_ba_n_step_05_newcc", "@InputMonth", ParaTime.Text.Trim(), app_configBA);
        }

        private void WcwChargeScope_Click(object sender, EventArgs e)
        {

        }

        private void CheckBUChanged_Click(object sender, EventArgs e)
        {
            string[] arr=null;
            String BuChanged = "";
            DataTable dt = ExcuteQuerySP("sp_ba_n_CheckOSMQBU", "", ParaTime.Text.Trim(), app_configBA);

            if (null == dt)
            {
                MessageBox.Show("get data fail!");
                return;
            }

            //dt.Rows
            FileStream fs =new FileStream(DefaultExcelPath + ParaTime.Text + check_BU+"temp.txt",FileMode.OpenOrCreate);

            StreamWriter sw = new StreamWriter(fs);




            foreach (DataRow dr in dt.Rows) {

                //(char[])("######")
                arr=dr[0].ToString().Split(new string[] { "######" }, StringSplitOptions.RemoveEmptyEntries);
                //processOSMQHistory(arr)
                sw.WriteLine(processOSMQHistory(arr));
                    
                //1 == arr;
            }

            sw.Close();
            fs.Close();


        }

        protected String processOSMQHistory(string[] arr) {
           // split 	######  arr1
            IList<String> SplitedList = new List<string>();
            String[] SplitDepartMent = null;
            //String OSMQ_split = "######";
            String OSMQ_BU_Tag = "Department:";

            String result = "";

            for (int i=1;i<arr.Length-1 ;i++ ) { 
            //arr[0] is not use
                if (arr[i].Contains(OSMQ_BU_Tag))

                    SplitDepartMent=arr[i].Split(new String[] { OSMQ_BU_Tag }, StringSplitOptions.RemoveEmptyEntries);

                SplitedList.Add( (getLastStrFromArr(SplitDepartMent[0].Trim().Split(new char[] {':'})).Trim() +"|"+ SplitDepartMent[1].Trim()) );              
            }

            foreach(String s in SplitedList){
                result = result + s +";"; 
            }

            return result;
        }


        public String getLastStrFromArr(String[] arr) {

            return (arr==null)? null:arr[arr.Length-1] ;            
        }
        
    }      
        
            
}
