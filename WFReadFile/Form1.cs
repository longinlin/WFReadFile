using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.Serialization;
using System.ComponentModel.DataAnnotations;
using Newtonsoft.Json;
using log4net;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Data.OleDb;
using System.Globalization;

namespace WFReadFile
{
    public partial class Form1 : Form
    {
        private static ILog log = LogManager.GetLogger(typeof(Form1)); //log4net
        private static string connectionString = "";
        public Form1()
        {
            connectionString = ConfigurationManager.AppSettings["ConnectionString"];
            InitializeComponent();
        }

        public class Consume
        {
            [DataMember]
            public string ConsumeType { get; set; }
            [DataMember]
            public string PatientClassCode { get; set; }
            [DataMember]
            public string AccountIdse { get; set; }
            [DataMember]
            public string Area { get; set; }
            [DataMember]
            public string PhrOrderIdse { get; set; }
            [DataMember]
            public string DrugCode { get; set; }
            [DataMember]
            public string Amount { get; set; }
            [DataMember]
            public string Device { get; set; }
            [DataMember]
            public string Empno { get; set; }
            [DataMember]
            public DateTime TransactionDateTime { get; set; }
        }


        public class OutStock
        {
            [DataMember]
            public string HospitalCode { get; set; }
            [DataMember]
            public string Area { get; set; }
            [DataMember]
            public string Device { get; set; }

            [DataMember]
            public string DrugCode { get; set; }
            [DataMember]
            public string Amount { get; set; }
            [DataMember]
            public DateTime TransactionDateTime { get; set; }
        }


        public string readjson(System.IO.StreamReader file, out bool IsSuccess)
        {
            StringBuilder JsonSB = new StringBuilder();
            string line = "";
            string jsonStr = "";
            IsSuccess = false;
            bool start = false, end = false;

            while ((line = file.ReadLine()) != null)
            {
                if (line == "{")
                {
                    start = true;
                    JsonSB.Append(line);
                    continue;
                }
                if (line == "}")
                {
                    JsonSB.Append(line);
                    end = true;
                    break;
                }
                if (start)
                    JsonSB.Append(line);

            }
            jsonStr = JsonSB.ToString();

            while ((line = file.ReadLine()) != null)
            {
                if (!string.IsNullOrEmpty(line))
                {
                    if (line.Contains("IsSuccess : False"))
                    {
                        IsSuccess = false;
                        break;
                    }
                    if (line.Contains("IsSuccess : True"))
                    {
                        IsSuccess = true;
                        break;
                    }
                }
            }

            return jsonStr;
        }

        public static DateTime StringToDateTime(string aDateStr)
        {
            string datetimestr = aDateStr;
            string[] DateTimeList = {
                    "yyyy/M/d tt hh:mm:ss",
                    "yyyy/MM/dd tt hh:mm:ss",
                    "yyyy/MM/dd HH:mm:ss",
                    "yyyy/M/d HH:mm:ss",
                    "yyyy/M/d",
                    "yyyy/MM/dd",
                    "yyyyMMdd",
                    "yyyyMMdd HHmm",
                    "yyyyMMdd HH:mm",
                    "yyyyMMdd HHmmss",
                    "yyyyMMddhhmmss tt",
                    "yyyyMMddhhmmsstt",
                    "yyyyMMddHHmmss",
                    "yyyy/MM/dd HH:mm",

                    "yyyy/M/d HH:mm",
                    "yyyy/M/d HH:m",
                    "yyyy/M/d H:mm",
                    "yyyy/M/d H:m",

                    "yyyy/M/dd HH:mm",
                    "yyyy/M/dd HH:m",
                    "yyyy/M/dd H:mm",
                    "yyyy/M/dd H:m",



                    "yyyy/M/dd",
                    "MM/dd/yyyy HH:mm",
                    "MMdd",
            };

            datetimestr = datetimestr.Replace("下午", "PM");
            datetimestr = datetimestr.Replace("上午", "AM");

            DateTime dt = DateTime.MinValue;
            try
            {
                dt = DateTime.ParseExact(datetimestr,
                DateTimeList,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AllowWhiteSpaces);
            }
            catch (Exception ex)
            {
                log.Error("StringToDateTime ADateStr:" + aDateStr);
            }
            return dt;

        }

        public static string getQueryFromCommand(SqlCommand cmd)
        {
            string CommandTxt = cmd.CommandText;
            Dictionary<string, string> Parameters = new Dictionary<string, string>();
            foreach (SqlParameter parms in cmd.Parameters)
            {
                Parameters.Add(parms.ParameterName, parms.ParameterName);
            }
            Parameters = Parameters.OrderByDescending(o => o.Key).ToDictionary(o => o.Key, p => p.Value);

            foreach (KeyValuePair<string, string> item in Parameters)
            {
                SqlParameter parms = cmd.Parameters[item.Key];
                string ParameterValue = String.Empty;
                if (parms.DbType.Equals(DbType.String) || parms.DbType.Equals(DbType.DateTime) || parms.DbType.Equals(DbType.AnsiString))
                {
                    ParameterValue = "'" + Convert.ToString(parms.Value).Replace(@"\", @"\\").Replace("'", @"\'") + "'";
                }
                else if (parms.DbType.Equals(DbType.Int16) || parms.DbType.Equals(DbType.Int32) || parms.DbType.Equals(DbType.Int64) || parms.DbType.Equals(DbType.Decimal) || parms.DbType.Equals(DbType.Double))
                {
                    ParameterValue = Convert.ToString(parms.Value);
                }
                CommandTxt = CommandTxt.Replace(parms.ParameterName, ParameterValue);
            }
            return (CommandTxt);
        }



        public DataTable TxtConvertToDataTable(string File, string TableName, string delimiter)
        {
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            StreamReader s = new StreamReader(File, System.Text.Encoding.Default);
            //string ss = s.ReadLine();//skip the first line
            string[] columns = s.ReadLine().Split(delimiter.ToCharArray());
            ds.Tables.Add(TableName);
            foreach (string col in columns)
            {
                bool added = false;
                string next = "";
                int i = 0;
                while (!added)
                {
                    string columnname = col + next;
                    columnname = columnname.Replace("#", "");
                    columnname = columnname.Replace("'", "");
                    columnname = columnname.Replace("&", "");

                    if (!ds.Tables[TableName].Columns.Contains(columnname))
                    {
                        ds.Tables[TableName].Columns.Add(columnname.ToUpper());
                        added = true;
                    }
                    else
                    {
                        i++;
                        next = "_" + i.ToString();
                    }
                }
            }

            string AllData = s.ReadToEnd();
            string[] rows = AllData.Split("\n".ToCharArray());

            foreach (string r in rows)
            {
                string[] items = r.Split(delimiter.ToCharArray());
                ds.Tables[TableName].Rows.Add(items);
            }

            s.Close();

            dt = ds.Tables[0];

            return dt;
        }

        public static DataTable ReadExcelAsTableNPOI(string fileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Open))
            {
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheetAt(0);
                DataTable table = new DataTable();
                //由第一列取標題做為欄位名稱
                IRow headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;
                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                    //以欄位文字為名新增欄位，此處全視為字串型別以求簡化
                    table.Columns.Add(
                        new DataColumn(headerRow.GetCell(i).StringCellValue));

                //略過第零列(標題列)，一直處理至最後一列
                for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;
                    DataRow dataRow = table.NewRow();
                    //依先前取得的欄位數逐一設定欄位內容
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                        if (row.GetCell(j) != null)
                            //如要針對不同型別做個別處理，可善用.CellType判斷型別
                            //再用.StringCellValue, .DateCellValue, .NumericCellValue...取值
                            //此處只簡單轉成字串
                            dataRow[j] = row.GetCell(j).ToString();
                    table.Rows.Add(dataRow);
                }
                return table;
            }
        }




        //DataTable dt = TxtConvertToDataTable(fileName, "tmp", ",");

        private void button1_Click(object sender, EventArgs e)
        {
            int counter = 0;
            string line;
            string error = "";
            // Read the file and display it line by line.  
            String PreLine = "";
            string LogFile = "";

            LogFile = txtBx_Log.Text;
            if (string.IsNullOrEmpty(LogFile))
            {
                MessageBox.Show("請選擇Log檔案");
                return;
            }
            string filename = txtBx_Log.Text;
            System.IO.StreamReader file =
                new System.IO.StreamReader(filename);
            //System.IO.StreamReader file =
            //    new System.IO.StreamReader(@"D:\Work\week\20200218\_Log20200218.txt");


            try
            {
                string[] truncatesql = { "truncate TABLE [Consume]", "truncate TABLE [NTUHData]", "truncate TABLE[OutStock]" };
                for (int i = 0; i < truncatesql.Count(); i++)
                {
                    using (SqlConnection conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        using (SqlCommand cmd = new SqlCommand(truncatesql[i], conn))
                        {
                            error = getQueryFromCommand(cmd);
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "truncate TABLE Error:" + error;
                msg += ex.Message;
            }

            while ((line = file.ReadLine()) != null)
            {
                if (line == "http://converws.ntuh.gov.tw/WebApplication/NTUHDrugWebAPI/ADC/Consume")
                {

                    //string aa = line;
                    string ajson = "";
                    bool IsSuccess = false;
                    ajson = readjson(file, out IsSuccess);
                    var aConsume = JsonConvert.DeserializeObject<Consume>(ajson);
                    aConsume.TransactionDateTime = StringToDateTime(PreLine);
                    try
                    {


                        string sql = "insert into Consume (ConsumeType,PatientClassCode,AccountIdse,Area,PhrOrderIdse,DrugCode,Amount,Device,Empno,TransactionDateTime,IsSuccess) values (@ConsumeType,@PatientClassCode,@AccountIdse,@Area,@PhrOrderIdse,@DrugCode,@Amount,@Device,@Empno,@TransactionDateTime,@IsSuccess)";
                        using (SqlConnection conn = new SqlConnection(connectionString))
                        {
                            conn.Open();
                            using (SqlCommand cmd = new SqlCommand(sql, conn))
                            {
                                //@PhrOrderIdse,@DrugCode,@Amount,@Device,@Empno,@TransactionDateTime
                                cmd.Parameters.Add("@ConsumeType", SqlDbType.VarChar);
                                cmd.Parameters["@ConsumeType"].Value = aConsume.ConsumeType;
                                cmd.Parameters.Add("@PatientClassCode", SqlDbType.VarChar);
                                cmd.Parameters["@PatientClassCode"].Value = aConsume.PatientClassCode;
                                cmd.Parameters.Add("@AccountIdse", SqlDbType.VarChar);
                                cmd.Parameters["@AccountIdse"].Value = aConsume.AccountIdse;
                                cmd.Parameters.Add("@Area", SqlDbType.VarChar);
                                cmd.Parameters["@Area"].Value = aConsume.Area;
                                cmd.Parameters.Add("@PhrOrderIdse", SqlDbType.VarChar);
                                cmd.Parameters["@PhrOrderIdse"].Value = aConsume.PhrOrderIdse;
                                cmd.Parameters.Add("@DrugCode", SqlDbType.VarChar);
                                cmd.Parameters["@DrugCode"].Value = aConsume.DrugCode;

                                Double DecAmount = 0;
                                if (!Double.TryParse(aConsume.Amount, out DecAmount))
                                {
                                    DecAmount = 0;
                                }
                                cmd.Parameters.Add("@Amount", SqlDbType.Decimal);
                                cmd.Parameters["@Amount"].Value = DecAmount;

                                cmd.Parameters.Add("@Device", SqlDbType.VarChar);
                                cmd.Parameters["@Device"].Value = aConsume.Device;

                                cmd.Parameters.Add("@Empno", SqlDbType.VarChar);
                                cmd.Parameters["@Empno"].Value = aConsume.Empno;

                                cmd.Parameters.Add("@TransactionDateTime", SqlDbType.DateTime);
                                cmd.Parameters["@TransactionDateTime"].Value = aConsume.TransactionDateTime;


                                cmd.Parameters.Add("@IsSuccess", SqlDbType.VarChar);
                                if (IsSuccess)
                                    cmd.Parameters["@IsSuccess"].Value = "Y";
                                else
                                    cmd.Parameters["@IsSuccess"].Value = "N";

                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    catch (SqlException ex)
                    {
                        string msg = "Insert Error:";
                        msg += ex.Message;
                    }

                    //MessageBox.Show(ajson);
                    //insert into Consume
                }
                if (line.Contains("http://converws.ntuh.gov.tw/WebApplication/NTUHDrugWebAPI/ADC/OutStock"))
                {
                    string ajson = "";
                    bool IsSuccess = false;
                    ajson = readjson(file, out IsSuccess);
                    var aOutStock = JsonConvert.DeserializeObject<OutStock>(ajson);
                    aOutStock.TransactionDateTime = StringToDateTime(PreLine);
                    //insert into aOutStock
                    try
                    {
                        string sql = "insert into OutStock (HospitalCode,Area,Device,DrugCode,Amount,TransactionDateTime,IsSuccess) values (@HospitalCode,@Area,@Device,@DrugCode,@Amount,@TransactionDateTime,@IsSuccess)";
                        using (SqlConnection conn = new SqlConnection(connectionString))
                        {
                            conn.Open();
                            using (SqlCommand cmd = new SqlCommand(sql, conn))
                            {
                                //@HospitalCode,@Area,@Device,@DrugCode,@Amount,@TransactionDateTime
                                cmd.Parameters.Add("@HospitalCode", SqlDbType.VarChar);
                                cmd.Parameters["@HospitalCode"].Value = aOutStock.HospitalCode;

                                cmd.Parameters.Add("@Area", SqlDbType.VarChar);
                                cmd.Parameters["@Area"].Value = aOutStock.Area;

                                cmd.Parameters.Add("@Device", SqlDbType.VarChar);
                                cmd.Parameters["@Device"].Value = aOutStock.Device;

                                cmd.Parameters.Add("@DrugCode", SqlDbType.VarChar);
                                cmd.Parameters["@DrugCode"].Value = aOutStock.DrugCode;



                                Decimal DecAmount = 0;
                                if (!Decimal.TryParse(aOutStock.Amount, out DecAmount))
                                {
                                    DecAmount = 0;
                                }
                                cmd.Parameters.Add("@Amount", SqlDbType.Decimal);
                                cmd.Parameters["@Amount"].Value = DecAmount;





                                cmd.Parameters.Add("@TransactionDateTime", SqlDbType.DateTime);
                                cmd.Parameters["@TransactionDateTime"].Value = aOutStock.TransactionDateTime;

                                cmd.Parameters.Add("@IsSuccess", SqlDbType.VarChar);
                                if (IsSuccess)
                                    cmd.Parameters["@IsSuccess"].Value = "Y";
                                else
                                    cmd.Parameters["@IsSuccess"].Value = "N";


                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    catch (SqlException ex)
                    {
                        string msg = "Insert Error:";
                        msg += ex.Message;
                    }
                }
                if (!string.IsNullOrEmpty(line))
                    PreLine = line;
                //System.Console.WriteLine(line);
                //  counter++;
            }

            file.Close();
            //System.Console.WriteLine("There were {0} lines.", counter);
            // Suspend the screen.  
            //System.Console.ReadLine();

            DataTable dt = new DataTable();
            DataTable dt2 = new DataTable();
            //dt = TxtConvertToDataTable(@"D:\Work\week\20200215\All Station Event Report - Sort by Station (Standard Size)NTUH -BIG5.csv", "NTUHData", ",");            
            //dt = ReadExcelAsTableNPOI(@"D:\Work\week\20200215\All Station Event Report - Sort by Station (Standard Size)NTUH -BIG5.xls");
            //dt = ReadExcelAsTableNPOI(@"D:\Work\week\20200215\All.xls");
            //dt = ReadExcelAsTableNPOI(@"D:\Work\week\20200215\All Station Event Report - Sort by Station (Standard Size)NTUH -BIG5.xls");
            //dt = ReadExcelAsTableNPOI(@"D:\Work\week\20202016\All Station Event Report - Sort by Station (Standard Size)_20200216.xls");
            //dt = ReadExcelAsTableNPOI(@"D:\Work\week\20200218\All Station Event Report - Sort by Station (Standard Size)_20200218.xls");
            string filename2 = "";
            filename = txtBx_Excel.Text;
            filename2 = txtBx_Csv2.Text;
            //dt = ReadExcelAsTableNPOI(filename);
            string DateStr = LogFile.Substring(LogFile.LastIndexOf("_Log")).Replace(".txt", "").Replace("_Log", "");
            //dt = ReadExcelAsTableNPOI(filename);
            //dt2 = ReadExcelAsTableNPOI(filename2);
            dt = GetDataTableFromCsv(filename2, true, DateStr);
            if(!string.IsNullOrEmpty(filename))
                dt2 = GetDataTableFromCsv(filename, true, DateStr);
            
            dt.Merge(dt2, false, MissingSchemaAction.Ignore);

            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    Lab_csv_count.Text = dt.Rows.Count.ToString() + "筆";
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (string.IsNullOrEmpty(dt.Rows[i][0].ToString()))
                            continue;
                        try
                        {
                            string sql = "insert into NTUHData ([Device],[Area],[DrwSubDrwPkt],[MedID],[EntID],[PISAltID],[CustomField1],[CustomField2],[CustomField3],[FacAltID],[MedDescription],[BrandName],[MedClass],[TransactionType],[TransactionDateTime],[TransactionElement],[Fractional],[Quantity],[QtyUOM],[Beg],[End],[BegEndUOM],[DiscrepancyQuantity],[DiscrepancyUOM],[ResolutionDatetime],[AutoResolved],[DiscrepancyResolutionDesc],[DiscrepancyReason],[ResolvedBy],[Reason],[UndocumentedWaste],[WasteReason],[CdcNotice],[CdcResponse],[InstructionText],[UserResponse],[CdcCatetory],[UserID],[UserName],[UserType],[WitnessID],[WitnessFullName],[Lot],[Serial],[ExpirationDate],[Order],[Route],[OrderStartDateAndTime],[OrderEndDateAndTime],[OrderingPhysician],[ID],[PatientName],[DateOfBirth],[RoomBed],[SessionType],[SessionBegin],[SessionEnd]) values (@Device,@Area,@DrwSubDrwPkt,@MedID,@EntID,@PISAltID,@CustomField1,@CustomField2,@CustomField3,@FacAltID,@MedDescription,@BrandName,@MedClass,@TransactionType,@TransactionDateTime,@TransactionElement,@Fractional,@Quantity,@QtyUOM,@Beg,@End,@BegEndUOM,@DiscrepancyQuantity,@DiscrepancyUOM,@ResolutionDatetime,@AutoResolved,@DiscrepancyResolutionDesc,@DiscrepancyReason,@ResolvedBy,@Reason,@UndocumentedWaste,@WasteReason,@CdcNotice,@CdcResponse,@InstructionText,@UserResponse,@CdcCatetory,@UserID,@UserName,@UserType,@WitnessID,@WitnessFullName,@Lot,@Serial,@ExpirationDate,@Order,@Route,@OrderStartDateAndTime,@OrderEndDateAndTime,@OrderingPhysician,@ID,@PatientName,@DateOfBirth,@RoomBed,@SessionType,@SessionBegin,@SessionEnd)";
                            using (SqlConnection conn = new SqlConnection(connectionString))
                            {
                                conn.Open();
                                using (SqlCommand cmd = new SqlCommand(sql, conn))
                                {
                                    //@CustomField2,@CustomField3,@FacAltID,@MedDescription,@BrandName,@MedClass,@TransactionType,@TransactionDateTime,@TransactionElement,@Fractional,@Quantity,@QtyUOM,@Beg,@End,@BegEndUOM,@DiscrepancyQuantity,@DiscrepancyUOM,@ResolutionDatetime,@AutoResolved,@DiscrepancyResolutionDesc,@DiscrepancyReason,@ResolvedBy,@Reason,@UndocumentedWaste,@WasteReason,@CdcNotice,@CdcResponse,@InstructionText,@UserResponse,@CdcCatetory,@UserID,@UserName,@UserType,@WitnessID,@WitnessFullName,@Lot,@Serial,@ExpirationDate,@Order,@Route,@OrderStartDateAndTime,@OrderEndDateAndTime,@OrderingPhysician,@ID,@PatientName,@DateOfBirth,@RoomBed,@SessionType,@SessionBegin,@SessionEnd
                                    cmd.Parameters.Add("@Device", SqlDbType.VarChar);
                                    cmd.Parameters["@Device"].Value = dt.Rows[i][0].ToString();

                                    cmd.Parameters.Add("@Area", SqlDbType.VarChar);
                                    cmd.Parameters["@Area"].Value = dt.Rows[i][1].ToString();

                                    cmd.Parameters.Add("@DrwSubDrwPkt", SqlDbType.VarChar);
                                    cmd.Parameters["@DrwSubDrwPkt"].Value = dt.Rows[i][2].ToString();

                                    cmd.Parameters.Add("@MedID", SqlDbType.VarChar);
                                    cmd.Parameters["@MedID"].Value = dt.Rows[i][3].ToString();

                                    cmd.Parameters.Add("@EntID", SqlDbType.VarChar);
                                    cmd.Parameters["@EntID"].Value = dt.Rows[i][4].ToString();

                                    cmd.Parameters.Add("@PISAltID", SqlDbType.VarChar);
                                    cmd.Parameters["@PISAltID"].Value = dt.Rows[i][5].ToString();

                                    cmd.Parameters.Add("@CustomField1", SqlDbType.VarChar);
                                    cmd.Parameters["@CustomField1"].Value = dt.Rows[i][6].ToString();

                                    cmd.Parameters.Add("@CustomField2", SqlDbType.VarChar);
                                    cmd.Parameters["@CustomField2"].Value = dt.Rows[i][7].ToString();

                                    cmd.Parameters.Add("@CustomField3", SqlDbType.VarChar);
                                    cmd.Parameters["@CustomField3"].Value = dt.Rows[i][8].ToString();

                                    cmd.Parameters.Add("@FacAltID", SqlDbType.VarChar);
                                    cmd.Parameters["@FacAltID"].Value = dt.Rows[i][9].ToString();

                                    cmd.Parameters.Add("@MedDescription", SqlDbType.VarChar);
                                    cmd.Parameters["@MedDescription"].Value = dt.Rows[i][10].ToString();

                                    cmd.Parameters.Add("@BrandName", SqlDbType.VarChar);
                                    cmd.Parameters["@BrandName"].Value = dt.Rows[i][11].ToString();

                                    cmd.Parameters.Add("@MedClass", SqlDbType.VarChar);
                                    cmd.Parameters["@MedClass"].Value = dt.Rows[i][12].ToString();

                                    cmd.Parameters.Add("@TransactionType", SqlDbType.VarChar);
                                    cmd.Parameters["@TransactionType"].Value = dt.Rows[i][13].ToString();

                                    cmd.Parameters.Add("@TransactionDateTime", SqlDbType.DateTime);
                                    cmd.Parameters["@TransactionDateTime"].Value = StringToDateTime(dt.Rows[i][14].ToString());

                                    cmd.Parameters.Add("@TransactionElement", SqlDbType.VarChar);
                                    cmd.Parameters["@TransactionElement"].Value = dt.Rows[i][15].ToString();

                                    cmd.Parameters.Add("@Fractional", SqlDbType.VarChar);
                                    cmd.Parameters["@Fractional"].Value = dt.Rows[i][16].ToString();

                                    /*
                                    Decimal DecQuantity = 0;
                                    if (!Decimal.TryParse(NtuhparaStr[17], out DecQuantity))
                                    {
                                        DecQuantity = 0;
                                    }
                                    */

                                    cmd.Parameters.Add("@Quantity", SqlDbType.VarChar);
                                    cmd.Parameters["@Quantity"].Value = dt.Rows[i][17].ToString();

                                    cmd.Parameters.Add("@QtyUOM", SqlDbType.VarChar);
                                    cmd.Parameters["@QtyUOM"].Value = dt.Rows[i][18].ToString();

                                    cmd.Parameters.Add("@Beg", SqlDbType.VarChar);
                                    cmd.Parameters["@Beg"].Value = dt.Rows[i][19].ToString();

                                    cmd.Parameters.Add("@End", SqlDbType.VarChar);
                                    cmd.Parameters["@End"].Value = dt.Rows[i][20].ToString();

                                    cmd.Parameters.Add("@BegEndUOM", SqlDbType.VarChar);
                                    cmd.Parameters["@BegEndUOM"].Value = dt.Rows[i][21].ToString();

                                    cmd.Parameters.Add("@DiscrepancyQuantity", SqlDbType.VarChar);
                                    cmd.Parameters["@DiscrepancyQuantity"].Value = dt.Rows[i][22].ToString();

                                    cmd.Parameters.Add("@DiscrepancyUOM", SqlDbType.VarChar);
                                    cmd.Parameters["@DiscrepancyUOM"].Value = dt.Rows[i][23].ToString();

                                    cmd.Parameters.Add("@ResolutionDatetime", SqlDbType.VarChar);
                                    cmd.Parameters["@ResolutionDatetime"].Value = dt.Rows[i][24].ToString();

                                    cmd.Parameters.Add("@AutoResolved", SqlDbType.VarChar);
                                    cmd.Parameters["@AutoResolved"].Value = dt.Rows[i][25].ToString();

                                    cmd.Parameters.Add("@DiscrepancyResolutionDesc", SqlDbType.VarChar);
                                    cmd.Parameters["@DiscrepancyResolutionDesc"].Value = dt.Rows[i][26].ToString();

                                    cmd.Parameters.Add("@DiscrepancyReason", SqlDbType.VarChar);
                                    cmd.Parameters["@DiscrepancyReason"].Value = dt.Rows[i][27].ToString();

                                    cmd.Parameters.Add("@ResolvedBy", SqlDbType.VarChar);
                                    cmd.Parameters["@ResolvedBy"].Value = dt.Rows[i][28].ToString();

                                    cmd.Parameters.Add("@Reason", SqlDbType.VarChar);
                                    cmd.Parameters["@Reason"].Value = dt.Rows[i][29].ToString();

                                    cmd.Parameters.Add("@UndocumentedWaste", SqlDbType.VarChar);
                                    cmd.Parameters["@UndocumentedWaste"].Value = dt.Rows[i][30].ToString();

                                    cmd.Parameters.Add("@WasteReason", SqlDbType.VarChar);
                                    cmd.Parameters["@WasteReason"].Value = dt.Rows[i][31].ToString();

                                    cmd.Parameters.Add("@CdcNotice", SqlDbType.VarChar);
                                    cmd.Parameters["@CdcNotice"].Value = dt.Rows[i][32].ToString();

                                    cmd.Parameters.Add("@CdcResponse", SqlDbType.VarChar);
                                    cmd.Parameters["@CdcResponse"].Value = dt.Rows[i][33].ToString();

                                    cmd.Parameters.Add("@InstructionText", SqlDbType.VarChar);
                                    cmd.Parameters["@InstructionText"].Value = dt.Rows[i][34].ToString();

                                    cmd.Parameters.Add("@UserResponse", SqlDbType.VarChar);
                                    cmd.Parameters["@UserResponse"].Value = dt.Rows[i][35].ToString();

                                    cmd.Parameters.Add("@CdcCatetory", SqlDbType.VarChar);
                                    cmd.Parameters["@CdcCatetory"].Value = dt.Rows[i][36].ToString();

                                    cmd.Parameters.Add("@UserID", SqlDbType.VarChar);
                                    cmd.Parameters["@UserID"].Value = dt.Rows[i][37].ToString();

                                    cmd.Parameters.Add("@UserName", SqlDbType.VarChar);
                                    cmd.Parameters["@UserName"].Value = dt.Rows[i][38].ToString();

                                    cmd.Parameters.Add("@UserType", SqlDbType.VarChar);
                                    cmd.Parameters["@UserType"].Value = dt.Rows[i][39].ToString();

                                    cmd.Parameters.Add("@WitnessID", SqlDbType.VarChar);
                                    cmd.Parameters["@WitnessID"].Value = dt.Rows[i][40].ToString();

                                    cmd.Parameters.Add("@WitnessFullName", SqlDbType.VarChar);
                                    cmd.Parameters["@WitnessFullName"].Value = dt.Rows[i][41].ToString();

                                    cmd.Parameters.Add("@Lot", SqlDbType.VarChar);
                                    cmd.Parameters["@Lot"].Value = dt.Rows[i][42].ToString();

                                    cmd.Parameters.Add("@Serial", SqlDbType.VarChar);
                                    cmd.Parameters["@Serial"].Value = dt.Rows[i][43].ToString();

                                    cmd.Parameters.Add("@ExpirationDate", SqlDbType.VarChar);
                                    cmd.Parameters["@ExpirationDate"].Value = dt.Rows[i][44].ToString();

                                    cmd.Parameters.Add("@Order", SqlDbType.VarChar);
                                    cmd.Parameters["@Order"].Value = dt.Rows[i][45].ToString();

                                    cmd.Parameters.Add("@Route", SqlDbType.VarChar);
                                    cmd.Parameters["@Route"].Value = dt.Rows[i][46].ToString();


                                    //cmd.Parameters.Add("@OrderStartDateAndTime", SqlDbType.VarChar);
                                    //if (string.IsNullOrEmpty(dt.Rows[i][47].ToString()))
                                    //    cmd.Parameters["@OrderStartDateAndTime"].Value = DBNull.Value;
                                    //else
                                    //    cmd.Parameters["@OrderStartDateAndTime"].Value = dt.Rows[i][47].ToString();


                                    //cmd.Parameters.Add("@OrderEndDateAndTime", SqlDbType.VarChar);
                                    //if (string.IsNullOrEmpty(dt.Rows[i][48].ToString()))
                                    //    cmd.Parameters["@OrderEndDateAndTime"].Value = DBNull.Value;
                                    //else
                                    //    cmd.Parameters["@OrderEndDateAndTime"].Value = dt.Rows[i][48].ToString();


                                    cmd.Parameters.Add("@OrderStartDateAndTime", SqlDbType.DateTime);
                                    if (string.IsNullOrEmpty(dt.Rows[i][47].ToString()))
                                        cmd.Parameters["@OrderStartDateAndTime"].Value = DBNull.Value;
                                    else
                                        cmd.Parameters["@OrderStartDateAndTime"].Value = StringToDateTime(dt.Rows[i][47].ToString());

                                    cmd.Parameters.Add("@OrderEndDateAndTime", SqlDbType.DateTime);
                                    if (string.IsNullOrEmpty(dt.Rows[i][48].ToString()))
                                        cmd.Parameters["@OrderEndDateAndTime"].Value = DBNull.Value;
                                    else
                                        cmd.Parameters["@OrderEndDateAndTime"].Value = StringToDateTime(dt.Rows[i][48].ToString());

                                    cmd.Parameters.Add("@OrderingPhysician", SqlDbType.VarChar);
                                    cmd.Parameters["@OrderingPhysician"].Value = dt.Rows[i][49].ToString();

                                    cmd.Parameters.Add("@ID", SqlDbType.VarChar);
                                    cmd.Parameters["@ID"].Value = dt.Rows[i][50].ToString();

                                    cmd.Parameters.Add("@PatientName", SqlDbType.VarChar);
                                    cmd.Parameters["@PatientName"].Value = dt.Rows[i][51].ToString();



                                    //cmd.Parameters.Add("@DateOfBirth", SqlDbType.VarChar);
                                    //if (string.IsNullOrEmpty(dt.Rows[i][52].ToString()))
                                    //    cmd.Parameters["@DateOfBirth"].Value = DBNull.Value;
                                    //else
                                    //    cmd.Parameters["@DateOfBirth"].Value = dt.Rows[i][52].ToString();

                                    cmd.Parameters.Add("@DateOfBirth", SqlDbType.Date);
                                    if (string.IsNullOrEmpty(dt.Rows[i][52].ToString()))
                                        cmd.Parameters["@DateOfBirth"].Value = DBNull.Value;
                                    else
                                        cmd.Parameters["@DateOfBirth"].Value = StringToDateTime(dt.Rows[i][52].ToString()).Date;

                                    cmd.Parameters.Add("@RoomBed", SqlDbType.VarChar);
                                    cmd.Parameters["@RoomBed"].Value = dt.Rows[i][53].ToString();

                                    cmd.Parameters.Add("@SessionType", SqlDbType.VarChar);
                                    cmd.Parameters["@SessionType"].Value = dt.Rows[i][54].ToString();

                                    //cmd.Parameters.Add("@SessionBegin", SqlDbType.VarChar);
                                    //if (string.IsNullOrEmpty(dt.Rows[i][55].ToString()))
                                    //    cmd.Parameters["@SessionBegin"].Value = DBNull.Value;
                                    //else
                                    //    cmd.Parameters["@SessionBegin"].Value = dt.Rows[i][55].ToString();

                                    //cmd.Parameters.Add("@SessionEnd", SqlDbType.VarChar);
                                    //if (string.IsNullOrEmpty(dt.Rows[i][56].ToString()))
                                    //    cmd.Parameters["@SessionEnd"].Value = DBNull.Value;
                                    //else
                                    //    cmd.Parameters["@SessionEnd"].Value = dt.Rows[i][56].ToString();


                                    cmd.Parameters.Add("@SessionBegin", SqlDbType.DateTime);
                                    if (string.IsNullOrEmpty(dt.Rows[i][55].ToString()))
                                        cmd.Parameters["@SessionBegin"].Value = DBNull.Value;
                                    else
                                        cmd.Parameters["@SessionBegin"].Value = StringToDateTime(dt.Rows[i][55].ToString());

                                    cmd.Parameters.Add("@SessionEnd", SqlDbType.DateTime);
                                    if (string.IsNullOrEmpty(dt.Rows[i][56].ToString()))
                                        cmd.Parameters["@SessionEnd"].Value = DBNull.Value;
                                    else
                                        cmd.Parameters["@SessionEnd"].Value = StringToDateTime(dt.Rows[i][56].ToString());




                                    error = getQueryFromCommand(cmd);
                                    cmd.ExecuteNonQuery();
                                }
                                conn.Close();
                            }
                        }
                        catch (SqlException ex)
                        {
                            string msg = "Insert Error:";
                            msg += ex.Message;
                            log.Error("Insert Error:" + error);
                        }
                    }
                }

            }
            /*
            if (dt2 != null)
            {
                if (dt2.Rows.Count > 0)
                {
                    Lab_csv_count2.Text = dt2.Rows.Count.ToString() + "筆";
                    for (int i = 0; i < dt2.Rows.Count; i++)
                    {
                        if (string.IsNullOrEmpty(dt2.Rows[i][0].ToString()))
                            continue;
                        try
                        {
                            string sql = "insert into NTUHData ([Device],[Area],[DrwSubDrwPkt],[MedID],[EntID],[PISAltID],[CustomField1],[CustomField2],[CustomField3],[FacAltID],[MedDescription],[BrandName],[MedClass],[TransactionType],[TransactionDateTime],[TransactionElement],[Fractional],[Quantity],[QtyUOM],[Beg],[End],[BegEndUOM],[DiscrepancyQuantity],[DiscrepancyUOM],[ResolutionDatetime],[AutoResolved],[DiscrepancyResolutionDesc],[DiscrepancyReason],[ResolvedBy],[Reason],[UndocumentedWaste],[WasteReason],[CdcNotice],[CdcResponse],[InstructionText],[UserResponse],[CdcCatetory],[UserID],[UserName],[UserType],[WitnessID],[WitnessFullName],[Lot],[Serial],[ExpirationDate],[Order],[Route],[OrderStartDateAndTime],[OrderEndDateAndTime],[OrderingPhysician],[ID],[PatientName],[DateOfBirth],[RoomBed],[SessionType],[SessionBegin],[SessionEnd]) values (@Device,@Area,@DrwSubDrwPkt,@MedID,@EntID,@PISAltID,@CustomField1,@CustomField2,@CustomField3,@FacAltID,@MedDescription,@BrandName,@MedClass,@TransactionType,@TransactionDateTime,@TransactionElement,@Fractional,@Quantity,@QtyUOM,@Beg,@End,@BegEndUOM,@DiscrepancyQuantity,@DiscrepancyUOM,@ResolutionDatetime,@AutoResolved,@DiscrepancyResolutionDesc,@DiscrepancyReason,@ResolvedBy,@Reason,@UndocumentedWaste,@WasteReason,@CdcNotice,@CdcResponse,@InstructionText,@UserResponse,@CdcCatetory,@UserID,@UserName,@UserType,@WitnessID,@WitnessFullName,@Lot,@Serial,@ExpirationDate,@Order,@Route,@OrderStartDateAndTime,@OrderEndDateAndTime,@OrderingPhysician,@ID,@PatientName,@DateOfBirth,@RoomBed,@SessionType,@SessionBegin,@SessionEnd)";
                            using (SqlConnection conn = new SqlConnection(connectionString))
                            {
                                conn.Open();
                                using (SqlCommand cmd = new SqlCommand(sql, conn))
                                {
                                    //@CustomField2,@CustomField3,@FacAltID,@MedDescription,@BrandName,@MedClass,@TransactionType,@TransactionDateTime,@TransactionElement,@Fractional,@Quantity,@QtyUOM,@Beg,@End,@BegEndUOM,@DiscrepancyQuantity,@DiscrepancyUOM,@ResolutionDatetime,@AutoResolved,@DiscrepancyResolutionDesc,@DiscrepancyReason,@ResolvedBy,@Reason,@UndocumentedWaste,@WasteReason,@CdcNotice,@CdcResponse,@InstructionText,@UserResponse,@CdcCatetory,@UserID,@UserName,@UserType,@WitnessID,@WitnessFullName,@Lot,@Serial,@ExpirationDate,@Order,@Route,@OrderStartDateAndTime,@OrderEndDateAndTime,@OrderingPhysician,@ID,@PatientName,@DateOfBirth,@RoomBed,@SessionType,@SessionBegin,@SessionEnd
                                    cmd.Parameters.Add("@Device", SqlDbType.VarChar);
                                    cmd.Parameters["@Device"].Value = dt2.Rows[i][0].ToString();

                                    cmd.Parameters.Add("@Area", SqlDbType.VarChar);
                                    cmd.Parameters["@Area"].Value = dt2.Rows[i][1].ToString();

                                    cmd.Parameters.Add("@DrwSubDrwPkt", SqlDbType.VarChar);
                                    cmd.Parameters["@DrwSubDrwPkt"].Value = dt2.Rows[i][2].ToString();

                                    cmd.Parameters.Add("@MedID", SqlDbType.VarChar);
                                    cmd.Parameters["@MedID"].Value = dt2.Rows[i][3].ToString();

                                    cmd.Parameters.Add("@EntID", SqlDbType.VarChar);
                                    cmd.Parameters["@EntID"].Value = dt2.Rows[i][4].ToString();

                                    cmd.Parameters.Add("@PISAltID", SqlDbType.VarChar);
                                    cmd.Parameters["@PISAltID"].Value = dt2.Rows[i][5].ToString();

                                    cmd.Parameters.Add("@CustomField1", SqlDbType.VarChar);
                                    cmd.Parameters["@CustomField1"].Value = dt2.Rows[i][6].ToString();

                                    cmd.Parameters.Add("@CustomField2", SqlDbType.VarChar);
                                    cmd.Parameters["@CustomField2"].Value = dt2.Rows[i][7].ToString();

                                    cmd.Parameters.Add("@CustomField3", SqlDbType.VarChar);
                                    cmd.Parameters["@CustomField3"].Value = dt2.Rows[i][8].ToString();

                                    cmd.Parameters.Add("@FacAltID", SqlDbType.VarChar);
                                    cmd.Parameters["@FacAltID"].Value = dt2.Rows[i][9].ToString();

                                    cmd.Parameters.Add("@MedDescription", SqlDbType.VarChar);
                                    cmd.Parameters["@MedDescription"].Value = dt2.Rows[i][10].ToString();

                                    cmd.Parameters.Add("@BrandName", SqlDbType.VarChar);
                                    cmd.Parameters["@BrandName"].Value = dt2.Rows[i][11].ToString();

                                    cmd.Parameters.Add("@MedClass", SqlDbType.VarChar);
                                    cmd.Parameters["@MedClass"].Value = dt2.Rows[i][12].ToString();

                                    cmd.Parameters.Add("@TransactionType", SqlDbType.VarChar);
                                    cmd.Parameters["@TransactionType"].Value = dt2.Rows[i][13].ToString();

                                    cmd.Parameters.Add("@TransactionDateTime", SqlDbType.DateTime);
                                    cmd.Parameters["@TransactionDateTime"].Value = StringToDateTime(dt2.Rows[i][14].ToString());

                                    cmd.Parameters.Add("@TransactionElement", SqlDbType.VarChar);
                                    cmd.Parameters["@TransactionElement"].Value = dt2.Rows[i][15].ToString();

                                    cmd.Parameters.Add("@Fractional", SqlDbType.VarChar);
                                    cmd.Parameters["@Fractional"].Value = dt2.Rows[i][16].ToString();


                                    cmd.Parameters.Add("@Quantity", SqlDbType.VarChar);
                                    cmd.Parameters["@Quantity"].Value = dt2.Rows[i][17].ToString();

                                    cmd.Parameters.Add("@QtyUOM", SqlDbType.VarChar);
                                    cmd.Parameters["@QtyUOM"].Value = dt2.Rows[i][18].ToString();

                                    cmd.Parameters.Add("@Beg", SqlDbType.VarChar);
                                    cmd.Parameters["@Beg"].Value = dt2.Rows[i][19].ToString();

                                    cmd.Parameters.Add("@End", SqlDbType.VarChar);
                                    cmd.Parameters["@End"].Value = dt2.Rows[i][20].ToString();

                                    cmd.Parameters.Add("@BegEndUOM", SqlDbType.VarChar);
                                    cmd.Parameters["@BegEndUOM"].Value = dt2.Rows[i][21].ToString();

                                    cmd.Parameters.Add("@DiscrepancyQuantity", SqlDbType.VarChar);
                                    cmd.Parameters["@DiscrepancyQuantity"].Value = dt2.Rows[i][22].ToString();

                                    cmd.Parameters.Add("@DiscrepancyUOM", SqlDbType.VarChar);
                                    cmd.Parameters["@DiscrepancyUOM"].Value = dt2.Rows[i][23].ToString();

                                    cmd.Parameters.Add("@ResolutionDatetime", SqlDbType.VarChar);
                                    cmd.Parameters["@ResolutionDatetime"].Value = dt2.Rows[i][24].ToString();

                                    cmd.Parameters.Add("@AutoResolved", SqlDbType.VarChar);
                                    cmd.Parameters["@AutoResolved"].Value = dt2.Rows[i][25].ToString();

                                    cmd.Parameters.Add("@DiscrepancyResolutionDesc", SqlDbType.VarChar);
                                    cmd.Parameters["@DiscrepancyResolutionDesc"].Value = dt2.Rows[i][26].ToString();

                                    cmd.Parameters.Add("@DiscrepancyReason", SqlDbType.VarChar);
                                    cmd.Parameters["@DiscrepancyReason"].Value = dt2.Rows[i][27].ToString();

                                    cmd.Parameters.Add("@ResolvedBy", SqlDbType.VarChar);
                                    cmd.Parameters["@ResolvedBy"].Value = dt2.Rows[i][28].ToString();

                                    cmd.Parameters.Add("@Reason", SqlDbType.VarChar);
                                    cmd.Parameters["@Reason"].Value = dt2.Rows[i][29].ToString();

                                    cmd.Parameters.Add("@UndocumentedWaste", SqlDbType.VarChar);
                                    cmd.Parameters["@UndocumentedWaste"].Value = dt2.Rows[i][30].ToString();

                                    cmd.Parameters.Add("@WasteReason", SqlDbType.VarChar);
                                    cmd.Parameters["@WasteReason"].Value = dt2.Rows[i][31].ToString();

                                    cmd.Parameters.Add("@CdcNotice", SqlDbType.VarChar);
                                    cmd.Parameters["@CdcNotice"].Value = dt2.Rows[i][32].ToString();

                                    cmd.Parameters.Add("@CdcResponse", SqlDbType.VarChar);
                                    cmd.Parameters["@CdcResponse"].Value = dt2.Rows[i][33].ToString();

                                    cmd.Parameters.Add("@InstructionText", SqlDbType.VarChar);
                                    cmd.Parameters["@InstructionText"].Value = dt2.Rows[i][34].ToString();

                                    cmd.Parameters.Add("@UserResponse", SqlDbType.VarChar);
                                    cmd.Parameters["@UserResponse"].Value = dt2.Rows[i][35].ToString();

                                    cmd.Parameters.Add("@CdcCatetory", SqlDbType.VarChar);
                                    cmd.Parameters["@CdcCatetory"].Value = dt2.Rows[i][36].ToString();

                                    cmd.Parameters.Add("@UserID", SqlDbType.VarChar);
                                    cmd.Parameters["@UserID"].Value = dt2.Rows[i][37].ToString();

                                    cmd.Parameters.Add("@UserName", SqlDbType.VarChar);
                                    cmd.Parameters["@UserName"].Value = dt2.Rows[i][38].ToString();

                                    cmd.Parameters.Add("@UserType", SqlDbType.VarChar);
                                    cmd.Parameters["@UserType"].Value = dt2.Rows[i][39].ToString();

                                    cmd.Parameters.Add("@WitnessID", SqlDbType.VarChar);
                                    cmd.Parameters["@WitnessID"].Value = dt2.Rows[i][40].ToString();

                                    cmd.Parameters.Add("@WitnessFullName", SqlDbType.VarChar);
                                    cmd.Parameters["@WitnessFullName"].Value = dt2.Rows[i][41].ToString();

                                    cmd.Parameters.Add("@Lot", SqlDbType.VarChar);
                                    cmd.Parameters["@Lot"].Value = dt2.Rows[i][42].ToString();

                                    cmd.Parameters.Add("@Serial", SqlDbType.VarChar);
                                    cmd.Parameters["@Serial"].Value = dt2.Rows[i][43].ToString();

                                    cmd.Parameters.Add("@ExpirationDate", SqlDbType.VarChar);
                                    cmd.Parameters["@ExpirationDate"].Value = dt2.Rows[i][44].ToString();

                                    cmd.Parameters.Add("@Order", SqlDbType.VarChar);
                                    cmd.Parameters["@Order"].Value = dt2.Rows[i][45].ToString();

                                    cmd.Parameters.Add("@Route", SqlDbType.VarChar);
                                    cmd.Parameters["@Route"].Value = dt2.Rows[i][46].ToString();


                                    //cmd.Parameters.Add("@OrderStartDateAndTime", SqlDbType.VarChar);
                                    //if (string.IsNullOrEmpty(dt.Rows[i][47].ToString()))
                                    //    cmd.Parameters["@OrderStartDateAndTime"].Value = DBNull.Value;
                                    //else
                                    //    cmd.Parameters["@OrderStartDateAndTime"].Value = dt.Rows[i][47].ToString();


                                    //cmd.Parameters.Add("@OrderEndDateAndTime", SqlDbType.VarChar);
                                    //if (string.IsNullOrEmpty(dt.Rows[i][48].ToString()))
                                    //    cmd.Parameters["@OrderEndDateAndTime"].Value = DBNull.Value;
                                    //else
                                    //    cmd.Parameters["@OrderEndDateAndTime"].Value = dt.Rows[i][48].ToString();


                                    cmd.Parameters.Add("@OrderStartDateAndTime", SqlDbType.DateTime);
                                    if (string.IsNullOrEmpty(dt2.Rows[i][47].ToString()))
                                        cmd.Parameters["@OrderStartDateAndTime"].Value = DBNull.Value;
                                    else
                                        cmd.Parameters["@OrderStartDateAndTime"].Value = StringToDateTime(dt2.Rows[i][47].ToString());

                                    cmd.Parameters.Add("@OrderEndDateAndTime", SqlDbType.DateTime);
                                    if (string.IsNullOrEmpty(dt2.Rows[i][48].ToString()))
                                        cmd.Parameters["@OrderEndDateAndTime"].Value = DBNull.Value;
                                    else
                                        cmd.Parameters["@OrderEndDateAndTime"].Value = StringToDateTime(dt2.Rows[i][48].ToString());

                                    cmd.Parameters.Add("@OrderingPhysician", SqlDbType.VarChar);
                                    cmd.Parameters["@OrderingPhysician"].Value = dt2.Rows[i][49].ToString();

                                    cmd.Parameters.Add("@ID", SqlDbType.VarChar);
                                    cmd.Parameters["@ID"].Value = dt2.Rows[i][50].ToString();

                                    cmd.Parameters.Add("@PatientName", SqlDbType.VarChar);
                                    cmd.Parameters["@PatientName"].Value = dt2.Rows[i][51].ToString();



                                    //cmd.Parameters.Add("@DateOfBirth", SqlDbType.VarChar);
                                    //if (string.IsNullOrEmpty(dt.Rows[i][52].ToString()))
                                    //    cmd.Parameters["@DateOfBirth"].Value = DBNull.Value;
                                    //else
                                    //    cmd.Parameters["@DateOfBirth"].Value = dt.Rows[i][52].ToString();

                                    cmd.Parameters.Add("@DateOfBirth", SqlDbType.Date);
                                    if (string.IsNullOrEmpty(dt2.Rows[i][52].ToString()))
                                        cmd.Parameters["@DateOfBirth"].Value = DBNull.Value;
                                    else
                                        cmd.Parameters["@DateOfBirth"].Value = StringToDateTime(dt2.Rows[i][52].ToString()).Date;

                                    cmd.Parameters.Add("@RoomBed", SqlDbType.VarChar);
                                    cmd.Parameters["@RoomBed"].Value = dt2.Rows[i][53].ToString();

                                    cmd.Parameters.Add("@SessionType", SqlDbType.VarChar);
                                    cmd.Parameters["@SessionType"].Value = dt2.Rows[i][54].ToString();

                                    //cmd.Parameters.Add("@SessionBegin", SqlDbType.VarChar);
                                    //if (string.IsNullOrEmpty(dt.Rows[i][55].ToString()))
                                    //    cmd.Parameters["@SessionBegin"].Value = DBNull.Value;
                                    //else
                                    //    cmd.Parameters["@SessionBegin"].Value = dt.Rows[i][55].ToString();

                                    //cmd.Parameters.Add("@SessionEnd", SqlDbType.VarChar);
                                    //if (string.IsNullOrEmpty(dt.Rows[i][56].ToString()))
                                    //    cmd.Parameters["@SessionEnd"].Value = DBNull.Value;
                                    //else
                                    //    cmd.Parameters["@SessionEnd"].Value = dt.Rows[i][56].ToString();


                                    cmd.Parameters.Add("@SessionBegin", SqlDbType.DateTime);
                                    if (string.IsNullOrEmpty(dt2.Rows[i][55].ToString()))
                                        cmd.Parameters["@SessionBegin"].Value = DBNull.Value;
                                    else
                                        cmd.Parameters["@SessionBegin"].Value = StringToDateTime(dt2.Rows[i][55].ToString());

                                    cmd.Parameters.Add("@SessionEnd", SqlDbType.DateTime);
                                    if (string.IsNullOrEmpty(dt2.Rows[i][56].ToString()))
                                        cmd.Parameters["@SessionEnd"].Value = DBNull.Value;
                                    else
                                        cmd.Parameters["@SessionEnd"].Value = StringToDateTime(dt2.Rows[i][56].ToString());




                                    error = getQueryFromCommand(cmd);
                                    cmd.ExecuteNonQuery();
                                }
                                conn.Close();
                            }
                        }
                        catch (SqlException ex)
                        {
                            string msg = "Insert Error:";
                            msg += ex.Message;
                            log.Error("Insert Error:" + error);
                        }
                    }
                }




            }
            */
            
            MessageBox.Show("資料建立完畢");
        }


        //
        private void button2_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt = GetDataTableFromCsv(@"D:\Work\week\20200220\All Station Event Report - Sort by Station (Standard Size)_20200220.csv", true, "20200224");
            if (dt != null)
            {
                Lab_csv_count.Text = dt.Rows.Count.ToString();
            }
            /*
            DataTable dataTable = new DataTable();
            string query = "select * from NTUHData where [TransactionType] in ('補充' , '卸載','卸載退出', '裝載','減少庫存') ";

            //string connString = @"your connection string here";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dataTable);
                    conn.Close();
                    da.Dispose();
                }
            }
            string query2 = "select * from OutStock ";
            DataTable dataTable2 = new DataTable();
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(query2, conn))
                {

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dataTable);
                    conn.Close();
                    da.Dispose();
                }
            }

            DataTable GetOutStockTb = new DataTable();
            if(dataTable!=null)
            {
                if (dataTable.Rows.Count > 0)
                {
                    foreach(DataRow row in dataTable.Rows)
                    {
                        //string Device,string Area,string MedID,string Quantity, DateTime TransactionDateTime
                        Double DecAmount = 0;
                        if (!Double.TryParse(row.Field<string>("Quantity"), out DecAmount))
                        {
                            DecAmount = 0;
                        }

                        GetOutStockTb = GetOutStock(row.Field<string>("Device"), row.Field<string>("Area"), row.Field<string>("MedID"), row.Field<string>("Quantity"), row.Field<DateTime>("TransactionDateTime"));
                        if (GetOutStockTb != null)
                        {
                            if (GetOutStockTb.Rows.Count > 0)
                            {
                                dataTable.Select("Device='" + row.Field<string>("Device") + "' and Area = '" + row.Field<string>("Area") + "' and MedID = '" + row.Field<string>("MedID") + "' and Quantity = " + DecAmount.ToString() + " and abs(DATEDIFF(MINUTE, TransactionDateTime, cONVERT('2020/2/12 16:52:00'),dATETIME))  < 4","");
                            }
                        }
                    }
                }
            }



            */
        }
        static DataTable GetDataTableFromCsv(string path, bool isFirstRowHeader, string DateStr)
        {
            DataTable dataTable = new DataTable();
            string header = isFirstRowHeader ? "Yes" : "No";

            string pathOnly = Path.GetDirectoryName(path);
            string fileName = Path.GetFileName(path);
            string newfileName = fileName;

            newfileName = newfileName.Substring(newfileName.LastIndexOf("_") + 1);
            System.IO.File.Copy(path, pathOnly + "\\" + newfileName);



            string sql = @"SELECT * FROM [" + newfileName + "]";
            //string sql = @"SELECT * FROM [" + fileName + "]";

            //; Extended properties = 'Character set=UTF8;'
            using (OleDbConnection connection = new OleDbConnection(
                      @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathOnly +
                      ";Extended Properties=\"Text;CharacterSet=65001;HDR=" + header + "\""))
            using (OleDbCommand command = new OleDbCommand(sql, connection))
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
            {

                //dataTable.Locale = CultureInfo.CurrentCulture;
                adapter.Fill(dataTable);

            }
            System.IO.File.Delete(pathOnly + "\\" + newfileName);
            //string exeDate = newfileName.Replace(".csv", "");
            string exeDate = DateStr;
            DateTime nextDT = StringToDateTime(exeDate).AddDays(1);
            DataView dv = dataTable.DefaultView;
            dv.RowFilter = "TransactionDateTime<= #" + nextDT.ToString("MM/dd/yyyy") + "#";
            dataTable = dv.ToTable();
            string aa = dataTable.Columns[12].DataType.ToString();
            return dataTable;
        }
        //a.Device = b.Device and a.Area = b.Area and a.MedID = b.DrugCode and cast(a.Quantity as decimal(5, 1)) = b.Amount and abs(DATEDIFF(MINUTE, b.TransactionDateTime, a.TransactionDateTime))  < 4

        public DataTable GetOutStock(string Device, string Area, string MedID, string Quantity, DateTime TransactionDateTime)
        {
            DataTable dataTable = new DataTable();
            string query = "select * from OutStock WHERE Device=@Device AND Area=@Area AND DrugCode = @MedID AND Amount = @Quantity AND abs(DATEDIFF(MINUTE, TransactionDateTime, @TransactionDateTime))  < 4 ";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {

                    cmd.Parameters.Add("@Device", SqlDbType.VarChar);
                    cmd.Parameters["@Device"].Value = Device;

                    cmd.Parameters.Add("@Area", SqlDbType.VarChar);
                    cmd.Parameters["@Area"].Value = Area;

                    cmd.Parameters.Add("@MedID", SqlDbType.VarChar);
                    cmd.Parameters["@MedID"].Value = MedID;




                    Double DecAmount = 0;
                    if (!Double.TryParse(Quantity, out DecAmount))
                    {
                        DecAmount = 0;
                    }

                    cmd.Parameters.Add("@Quantity", SqlDbType.Decimal);
                    cmd.Parameters["@Quantity"].Value = DecAmount;

                    cmd.Parameters.Add("@TransactionDateTime", SqlDbType.DateTime);
                    cmd.Parameters["@TransactionDateTime"].Value = TransactionDateTime;

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dataTable);
                    conn.Close();
                    da.Dispose();
                }
                return dataTable;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string error = "";
            for (int i = 3; i < 14; i++)
            {
                StringBuilder sb1 = new StringBuilder();
                StringBuilder sb2 = new StringBuilder();
                try
                {

                    // update NTUHData
                    sb1.Append("update NTUHData set NTUHData.MappingID = d.ConsumeID ");
                    sb1.Append("from NTUHData c, ");
                    sb1.Append(" ( ");
                    sb1.Append(" select ROW_NUMBER() OVER(ORDER BY a.NTUHDataID) AS rownum, a.NTUHDataID,b.ConsumeID ");
                    sb1.Append(" from NTUHData a inner join Consume b on ");
                    sb1.Append(" a.Device = b.Device ");
                    sb1.Append(" and a.Area = b.Area ");
                    sb1.Append(" and a.MedID = b.DrugCode ");
                    sb1.Append(" and ABS(DATEDIFF(MINUTE, a.TransactionDateTime, b.TransactionDateTime)) <=  " + i.ToString() + " ");
                    sb1.Append(" and cast(a.Quantity as decimal(5, 1)) = b.Amount ");
                    sb1.Append(" and (a.UserID = b.Empno  or (CASE WHEN ISNUMERIC(b.Empno) = 1 THEN  case when(cast(a.UserID as int) - cast(b.Empno as int)) = 0 then 1 else 0 end ELSE 0 END) = 1) ");
                    //sb1.Append(" and (a.UserID = b.Empno  or  cast(a.UserID as int) = cast(b.Empno as int)  )");
                    sb1.Append(" and a.[Order] = b.PhrOrderIdse ");
                    sb1.Append(" where  a.TransactionType in ('移除','應急取藥') ");
                    sb1.Append(" and a.MappingID is null and b.MappingID is null ");
                    sb1.Append(" ) d where c.NTUHDataID = d.NTUHDataID and d.rownum = 1 ");

                    // update Consume
                    sb2.Append(" update Consume ");
                    sb2.Append(" set Consume.MappingID = a.NTUHDataID ");
                    sb2.Append(" from NTUHData a INNER ");
                    sb2.Append(" join Consume b ");
                    sb2.Append(" ON ");
                    sb2.Append(" a.Device = b.Device ");
                    sb2.Append(" and a.Area = b.Area ");
                    sb2.Append(" and a.MedID = b.DrugCode ");
                    sb2.Append(" and ABS(DATEDIFF(MINUTE, a.TransactionDateTime, b.TransactionDateTime)) <=  " + i.ToString() + " ");
                    sb2.Append(" and cast(a.Quantity as decimal(5, 1)) = b.Amount ");
                    sb2.Append(" and (a.UserID = b.Empno  or (CASE WHEN ISNUMERIC(b.Empno) = 1 THEN  case when(cast(a.UserID as int) - cast(b.Empno as int)) = 0 then 1 else 0 end ELSE 0 END) = 1) ");
                    //sb2.Append(" and (a.UserID = b.Empno or  cast(a.UserID as int) = cast(b.Empno as int)) ");
                    sb2.Append(" and a.[Order] = b.PhrOrderIdse ");
                    sb2.Append(" and a.MappingID  = b.ConsumeID ");
                    sb2.Append(" where b.MappingID is null ");

                    string sql1 = sb1.ToString();
                    string sql2 = sb2.ToString();
                    int effrow = 0;
                    while (true)
                    {
                        using (SqlConnection conn = new SqlConnection(connectionString))
                        {
                            conn.Open();
                            using (SqlCommand cmd = new SqlCommand(sql1, conn))
                            {
                                error = getQueryFromCommand(cmd);
                                effrow = cmd.ExecuteNonQuery();
                                if (effrow == 0)
                                    break;
                            }
                            using (SqlCommand cmd = new SqlCommand(sql2, conn))
                            {
                                error = getQueryFromCommand(cmd);
                                effrow = cmd.ExecuteNonQuery();
                            }
                            conn.Close();
                        }
                    }


                    //NTUHData
                    sb1 = new StringBuilder();
                    sb1.Append(" update NTUHData set NTUHData.OutStockID = d.OutStockID ");
                    sb1.Append(" from NTUHData c, ");
                    sb1.Append(" ( ");
                    sb1.Append(" select ROW_NUMBER() OVER(ORDER BY NTUHDataID) AS rownum, a.NTUHDataID, b.OutStockID ");
                    sb1.Append(" from NTUHData a INNER ");
                    sb1.Append(" join OutStock b ");
                    sb1.Append(" ON a.Device = b.Device and a.Area = b.Area and a.MedID = b.DrugCode and cast(a.Quantity as decimal(5, 1)) = ");
                    sb1.Append(" ((case when a.TransactionType in('卸載','卸載退出') then  -1 else case when a.TransactionType='減少庫存' then -1 else 1 end  end) * b.Amount)  and abs(DATEDIFF(MINUTE, b.TransactionDateTime, a.TransactionDateTime)) <=  " + i.ToString() + " ");
                    sb1.Append(" where ");
                    sb1.Append(" a.TransactionType in ('補充' , '卸載','卸載退出', '裝載', '減少庫存') ");
                    sb1.Append(" and a.OutStockID is null and b.MappingID is null ");
                    sb1.Append(" ) d where c.NTUHDataID = d.NTUHDataID and d.rownum = 1 ");
                    // update Consume
                    sb2 = new StringBuilder();

                    sb2.Append(" update OutStock set OutStock.MappingID = a.NTUHDataID ");
                    sb2.Append(" from NTUHData a INNER ");
                    sb2.Append(" join OutStock b ");
                    sb2.Append(" ON a.Device = b.Device and a.Area = b.Area and a.MedID = b.DrugCode and cast(a.Quantity as decimal(5, 1)) = ((case when a.TransactionType in('卸載','卸載退出') then -1 else case when a.TransactionType='減少庫存' then -1 else 1 end end) * b.Amount)  and abs(DATEDIFF(MINUTE, b.TransactionDateTime, a.TransactionDateTime)) <= " + i.ToString() + " ");
                    sb2.Append(" and a.OutStockID = b.OutStockID ");
                    sb2.Append(" where b.MappingID is null ");

                    sql1 = sb1.ToString();
                    sql2 = sb2.ToString();


                    effrow = 0;
                    while (true)
                    {
                        using (SqlConnection conn = new SqlConnection(connectionString))
                        {
                            conn.Open();
                            using (SqlCommand cmd = new SqlCommand(sql1, conn))
                            {
                                error = getQueryFromCommand(cmd);
                                effrow = cmd.ExecuteNonQuery();
                                if (effrow == 0)
                                    break;
                            }
                            using (SqlCommand cmd = new SqlCommand(sql2, conn))
                            {
                                error = getQueryFromCommand(cmd);
                                effrow = cmd.ExecuteNonQuery();
                            }
                            conn.Close();
                        }
                    }
                }
                catch (SqlException ex)
                {
                    string msg = "Insert Error:";
                    msg += ex.Message;
                    log.Error("Insert Error:" + error);
                }
            }
            MessageBox.Show("資料比對完成");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;//該值確定是否可以選擇多個檔案
            dialog.Title = "請選擇資料夾";
            dialog.Filter = "所有檔案(*.txt)|*.txt";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string file = dialog.FileName;
                txtBx_Log.Text = file;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {



            StringBuilder sb = new StringBuilder();
            string[] querystr = new string[3];
            //sql1
            sb.Append(" SELECT [NTUHDataid] ");
            sb.Append(" ,[Device],[Area] ,[MedID],DrwSubDrwPkt,[MedClass],[TransactionType],[TransactionDateTime] ");
            sb.Append(" ,[TransactionElement],[Fractional],[Quantity],[UserID],[UserName],[Order],[Route],[OrderStartDateAndTime],[OrderEndDateAndTime] ");
            sb.Append(" ,[OrderingPhysician],[ID],[PatientName],[DateOfBirth],[RoomBed],[SessionType],[SessionBegin],[SessionEnd] ");
            sb.Append(" FROM[PyxisCheck].[dbo].[NTUHData] ");
            sb.Append(" where TransactionType in ('補充' , '卸載','卸載退出', '裝載','減少庫存','移除','應急取藥') ");
            sb.Append(" and OutStockId is null and MappingID is null ");
            sb.Append(" and [Quantity] <> '0' ");
            sb.Append(" and [TransactionElement] <> '完整' ");

            querystr[0] = sb.ToString();

            sb = new StringBuilder();

            sb.Append(" SELECT b.IsSuccess OutStockSuccess ");
            sb.Append(" , a.[Device], a.[Area] ,a.[MedID] ,a.[MedClass] ,a.[TransactionType] ");
            sb.Append(" , a.[TransactionDateTime], a.[TransactionElement] ,a.[Fractional] ,a.[Quantity] ,a.[UserID] ");
            sb.Append(" ,a.[UserName] , a.[Order] ,a.[Route] ,a.[OrderStartDateAndTime] ,a.[OrderEndDateAndTime] ");
            sb.Append(" ,a.[OrderingPhysician], a.[ID] ,a.[PatientName] ,a.[DateOfBirth] ,a.[RoomBed] ");
            sb.Append(" ,a.[SessionType], a.[SessionBegin] ,a.[SessionEnd] ");
            sb.Append(" FROM[PyxisCheck].[dbo].[NTUHData] ");
            sb.Append(" a, OutStock b ");
            sb.Append(" where a.TransactionType in ('補充' , '卸載','卸載退出', '裝載','減少庫存','移除','應急取藥') ");
            sb.Append(" and a.[Quantity] <> '0' ");
            sb.Append(" and a.[TransactionElement] <> '完整' ");
            sb.Append(" and a.OutStockID = b.OutStockID ");
            sb.Append(" and b.IsSuccess = 'N' ");

            querystr[1] = sb.ToString();
            sb = new StringBuilder();

            sb.Append(" SELECT b.IsSuccess ConsumeSuccess, ");
            sb.Append(" a.[Device], a.[Area] ,a.[MedID],a.[MedClass],a.[TransactionType] ");
            sb.Append(" ,a.[TransactionDateTime], a.[TransactionElement],a.[Fractional],a.[Quantity] ,a.[UserID] ");
            sb.Append(" ,a.[UserName] , a.[Order] ,a.[Route],a.[OrderStartDateAndTime],a.[OrderEndDateAndTime] ");
            sb.Append(" ,a.[OrderingPhysician], a.[ID],a.[PatientName],a.[DateOfBirth] ,a.[RoomBed] ");
            sb.Append(" ,a.[SessionType], a.[SessionBegin],a.[SessionEnd] ");
            sb.Append(" FROM[PyxisCheck].[dbo].[NTUHData] ");
            sb.Append(" a, Consume b ");
            sb.Append(" where a.TransactionType in ('補充' , '卸載','卸載退出', '裝載','減少庫存','移除','應急取藥') ");
            sb.Append(" and a.[Quantity] <> '0' ");
            sb.Append(" and a.[TransactionElement] <> '完整' ");
            sb.Append(" and a.mappingid = b.ConsumeID ");
            sb.Append(" and b.IsSuccess = 'N' ");
            querystr[2] = sb.ToString();

            DataSet DS = new DataSet();
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(querystr[0], conn))
                {

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(DS, "未對應資料");
                    conn.Close();
                    da.Dispose();
                }
            }

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(querystr[1], conn))
                {

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(DS, "OutStock上傳異常");
                    conn.Close();
                    da.Dispose();
                }
            }

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(querystr[2], conn))
                {

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(DS, "Consume上傳異常");
                    conn.Close();
                    da.Dispose();
                }
            }
            //string filename = @"D:\Work\week\20200218\_Log20200219.txt";
            string filename = txtBx_Log.Text;
            int index = 0;
            string datestr = "";
            index = filename.LastIndexOf('\\');
            datestr = filename.Substring(index + 1).Replace("_Log", "").Replace(".txt", "");
            filename = filename.Substring(0, index) + "\\";


            filename = filename + "異常資料" + datestr + ".xlsx";
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }

            ExportToExcel.CreateExcelFile.CreateExcelDocument(DS, filename);
            MessageBox.Show("產生異常報表完成 路徑[" + filename + "]");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;//該值確定是否可以選擇多個檔案
            dialog.Title = "請選擇資料夾";
            dialog.Filter = "所有檔案(*.csv)|*.csv";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string file = dialog.FileName;
                txtBx_Excel.Text = file;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;//該值確定是否可以選擇多個檔案
            dialog.Title = "請選擇資料夾";
            dialog.Filter = "所有檔案(*.csv)|*.csv";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string file = dialog.FileName;
                txtBx_Csv2.Text = file;
            }
        }


        public System.Data.DataTable GetDataTable(string FilePath)
        {
            #region Local Member Variables
            string folderPath = System.IO.Path.GetDirectoryName(FilePath);
            string fileName = System.IO.Path.GetFileName(FilePath);
            System.IO.FileInfo Info = new System.IO.FileInfo(FilePath);
            System.Data.DataTable ds = new System.Data.DataTable();
            #endregion

            //
            if (Info.Extension.ToLower() == ".csv")
            {
                System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection
                ("Provider=Microsoft.Jet.OleDb.4.0; Data Source = " + folderPath + "; Extended Properties = \"Text;CharacterSet=65001;HDR=YES;FMT=Delimited\"");
                /*
                using (OleDbConnection connection = new OleDbConnection(
                   @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathOnly +
                   ";Extended Properties=\"Text;CharacterSet=UTF8;IMEX=1;TypeGuessRows=2;ImportMixedTypes=Text;HDR=" + header + "\""))
                   */
                // Opens a database connection
                conn.Open();

                //SELECT *, cast(PostalCode as nvarchar) as PostalCode FROM [" + fileName + "]
                string strQuery = @"select str(MedClass) as MedClass from [" + fileName + "]";

                System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter(strQuery, conn);

                // Adds or refreshes rows in the dataset
                adapter.Fill(ds);

                // Closes the connection
                conn.Close();
            }
            return ds;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            //dt = GetDataTable(@"D:\Work\week\20200224\NTUH\NTUH_0221.csv");
            dt = ImportCSVData(@"D:\Work\week\20200224\NTUH\NTUH_0221.csv");

        }

        private DataTable ImportCSVData(string strCSVFilePath)
        {
            string folderPath = System.IO.Path.GetDirectoryName(strCSVFilePath);
            string fileName = System.IO.Path.GetFileName(strCSVFilePath);

            DataTable dt = new DataTable();
            OleDbConnection conn = null;
            try
            {


                string strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + folderPath + ";Extended Properties='text;HDR=Yes;FMT=Delimited(,)';";
                string sql_select;
                conn = new OleDbConnection(strConnString.Trim());
                sql_select = "select * from [" + fileName + "]";
                conn.Open();
                OleDbCommand cmd = new OleDbCommand(sql_select, conn);


                OleDbDataAdapter obj_oledb_da = new OleDbDataAdapter(cmd);
                DataTable dtSchema = new DataTable();
                obj_oledb_da.FillSchema(dtSchema, SchemaType.Source);


                if (dtSchema != null)
                    writeSchema(dtSchema,folderPath,fileName);


                obj_oledb_da.Fill(dt);




            }
            finally
            {
                if (conn.State == System.Data.ConnectionState.Open)
                    conn.Close();
            }
            return dt;
        }




        private void writeSchema(DataTable dt, string folderPath, string fileName)
        {
            try
            {
                FileStream fsOutput = new FileStream(folderPath + "\\schema.ini", FileMode.Create, FileAccess.Write);
                StreamWriter srOutput = new StreamWriter(fsOutput);
                string s1, s2, s3, s4, s5,s6;
                s1 = "[" + fileName + "]";
                s2 = "ColNameHeader=True";
                s3 = "Format=CSVDelimited";
                s4 = "MaxScanRows=50";
                s5 = "CharacterSet=ANSI";
                //s6 = "Col1=MedClass Text";
                //s6 = "DateTimeFormat=MM/dd/yyyy";
                srOutput.WriteLine(s1 + '\n' + s2 + '\n' + s3 + '\n' + s4 + '\n' + s5);
                StringBuilder strB = new StringBuilder();
                if (dt != null)
                {
                    for (Int32 ColIndex = 1; ColIndex <= dt.Columns.Count; ColIndex++)
                    {
                        strB.Append("Col" + ColIndex.ToString());
                        strB.Append("=F" + ColIndex.ToString());
                        strB.Append(" Text\n");
                        srOutput.WriteLine(strB.ToString());
                        strB = new StringBuilder();
                    }
                }


                srOutput.Close();
                fsOutput.Close();
            }
            catch (Exception ex)
            {


                log.Info("Exception", ex);
            }
        }
    }

}
