using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SqlServerHelper.Core;
using SqlServerHelper;
using System.ComponentModel;
using LicenseContext = OfficeOpenXml.LicenseContext;
using System.Reflection;
using System.ComponentModel.DataAnnotations;

namespace SchedulerDB
{
    public static class ExcelExtensions
    {
        // SetQuickStyle，指定前景色/背景色/水平對齊
        public static void SetQuickStyle(this ExcelRange range,
            Color fontColor,
            Color bgColor = default(Color),
            ExcelHorizontalAlignment hAlign = ExcelHorizontalAlignment.Left)
        {
            range.Style.Font.Color.SetColor(fontColor);
            if (bgColor != default(Color))
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid; // 一定要加這行..不然會報錯
                range.Style.Fill.BackgroundColor.SetColor(bgColor);
            }
            range.Style.HorizontalAlignment = hAlign;
        }

        //讓文字上有連結
        public static void SetHyperlink(this ExcelRange range, Uri uri)
        {
            range.Hyperlink = uri;
            range.Style.Font.UnderLine = true;
            range.Style.Font.Color.SetColor(Color.Blue);
        }
    }
    class Program
    {

        static void Main(string[] args)
        {
            string sql = @"IF object_id('tempdb..#RESTTReservation') IS NOT NULL DROP TABLE #RESTTReservation

                            SELECT a.RoomCode RESRoomCode,
                                   b.RoomCode XRYRoomCode,a.CalendarGroupName,
                                   a.MedicalNoteNo,a.ExaRequestNo,a.[Start],b.PlanDate ,a.ReservationSourceType,a.SourceCode,
                                   b.DVC_CHRT,b.DVC_RQNO,b.DVC_DATE,b.DVC_STTM,b.SourceCode XRYSourceCode
                            INTO #RESTTReservation
                            FROM 
                            (
                              SELECT d.RoomCode,d.CalendarGroupName,a.MedicalNoteNo,a.ExaRequestNo,c.[Start],a.ReservationSourceType,
                                     'RESTTReservation' SourceCode
                              FROM HISSCHDB.dbo.RESTTReservation a
                              INNER JOIN HISSCHDB.dbo.RESTTimeslotRes b ON a.Id = b.ReservationId
                              INNER JOIN HISSCHDB.dbo.RESTTimeslot c ON b.TimeslotId = c.Id
                              LEFT JOIN HISSCHDB.dbo.tmpEXAMRoomMapping d ON a.CalendarId = d.CalendarId
                              WHERE c.[Start] >= convert(date,dateadd(day,-30,getdate()))
                              --ORDER BY d.RoomCode,c.[Start]
                            ) a
                            FULL OUTER JOIN 
                            (
                             SELECT SUBSTRING(a.DVC_ROOM,2,2) RoomCode,a.DVC_CHRT,a.DVC_RQNO,

                            CASE WHEN ISNUMERIC(DVC_DATE) = 1 AND ISNUMERIC(DVC_STTM) = 1 
                            THEN try_convert(datetime,convert(varchar, substring((
                                            (CASE WHEN LEN(rtrim(DVC_DATE)) = 7 
                                                  THEN DVC_DATE 
                             ELSE (CASE WHEN LEFT(DVC_DATE,1) = '0' THEN '1' ELSE '0' END) + DVC_DATE END )) ,1,3)+ 1911)+RIGHT(DVC_DATE,4) +' '+substring(DVC_STTM,1,2)+':'+substring(DVC_STTM,3,2))     
                             WHEN ISNUMERIC(DVC_DATE) = 1 
                             THEN try_convert(date,convert(varchar, substring((
                             (CASE WHEN LEN(rtrim(DVC_DATE)) = 7 
                                   THEN DVC_DATE 
                                ELSE (CASE WHEN LEFT(DVC_DATE,1) = '0' THEN '1' ELSE '0' END) + DVC_DATE END )),1,3)+ 1911)+RIGHT(DVC_DATE,4))     
                             END PlanDate,
                              a.DVC_DATE,a.DVC_STTM,'XRYMDVCF' SourceCode
                              FROM [10.1.222.182].PXRYDB.SKDBA.XRYMDVCF a
                              WHERE rtrim(DVC_CHRT) <> ''
                              AND rtrim(DVC_RQNO) <> ''
                              AND a.DVC_DATE >=  convert(varchar,dateadd(day,-30,getdate()),112)-19110000

                             ) b ON a.MedicalNoteNo = b.DVC_CHRT AND a.ExaRequestNo = b.DVC_RQNO

                             select RESRoomCode,XRYRoomCode,CalendarGroupName,MedicalNoteNo,ExaRequestNo,Start,DVC_CHRT,DVC_RQNO,DVC_DATE,DVC_STTM from #RESTTReservation order by RESRoomCode;
                                 ";

            //Step 1.讀取DB Table List
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsetting.json", optional: true, reloadOnChange: true).Build();
            //取得連線字串
            string connString = config.GetConnectionString("DefaultConnection");
            //string connString = "Data Source=10.1.222.181;Initial Catalog={0};Integrated Security=False;User ID={1};Password={2};Pooling=True;MultipleActiveResultSets=True;Connect Timeout=120;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite";
            SqlServerDBHelper sqlHelper = new SqlServerDBHelper(string.Format(connString, "HISDB", "msdba", "1qaz@wsx"));
            // DataTable dt = sqlHelper.FillTableAsync(sql).Result;
            DataTable dt = sqlHelper.FillTableAsync(sql).Result;
            //印出list
            //int rowCount = (dt == null) ? 0 : dt.Rows.Count;
            //Console.WriteLine(rowCount);

            //Step 1.1.將資料放入List
            List<DBData> migrationTableInfoList = sqlHelper.QueryAsync<DBData>(sql).Result?.ToList();
            //Step 1.2 將date Distinct排序給sheet用 > 遞增 order by 遞減OrderByDescending
            var datetime = migrationTableInfoList.Select(p => p.Start != DateTime.MinValue ? p.Start.Date : p.PlanDate.Date)
                                                 .OrderBy(p => p.Date)
                                                 .Distinct()
                                                 .ToList();
            //Step 2.建立 各日期Sheet
            // var excelname = "Scheduler" + DateTime.Now.ToString("yyyyMMddhhmm") + ".xlsx";
            var excelname = new FileInfo("Scheduler" + DateTime.Now.ToString("yyyyMMddhhmm") + ".xlsx");
            //ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excel = new ExcelPackage(excelname))
            {
                var importDBData = new ImportDBData();
                importDBData.GenFirstSheet(excel, datetime);
                for (int sheetnum = 0; sheetnum <= datetime.Count - 1; sheetnum++)
                {
                    //Step 3.將對應的List 丟到各Sheet中
                    ExcelWorksheet sheet = excel.Workbook.Worksheets.Add(datetime[sheetnum].ToString("yyyy-MM-dd"));
                    //抽function
                    int rowIndex = 2;
                    int colIndex = 1;
                    importDBData.ImportData(dt, sheet, rowIndex, colIndex, migrationTableInfoList);
                }
                // Step 4.Export EXCEL
                Byte[] bin = excel.GetAsByteArray();
                File.WriteAllBytes(@"D:\微軟MCS\SchedulerDB_Excel\" + excelname, bin);

            }

            //Step 5. Send Email
            var helper = new SMTPHelper("lovemath0630@gmail.com", "koormyktfbbacpmj", "smtp.gmail.com", 587, true, true); //寄出信email
            string subject = $"Datebase Scheduler報表 {DateTime.Now.ToString("yyyyMMdd")}"; //信件主旨
            string body = $"Hi All, \r\n\r\n{DateTime.Now.ToString("yyyyMMdd")} Scheduler報表 如附件，\r\n\r\n Best Regards, \r\n\r\n Vicky Yin";//信件內容
            string attachments = null;//附件
            var fileName = @"D:\微軟MCS\SchedulerDB_Excel\" + excelname;//附件位置
            if (File.Exists(fileName.ToString()))
            {
                attachments = fileName.ToString();
            }
            string toMailList = "lovemath0630@gmail.com;v-vyin@microsoft.com";//收件者
            string ccMailList = "";//CC收件者

            helper.SendMail(toMailList, ccMailList, null, subject, body, attachments);
        }

        public class DBData
        {
            [Required]
            [DisplayName("新檢查室")]
            public string RESRoomCode { get; set; }
            [Required]
            [DisplayName("舊檢查室")]
            public string XRYRoomCode { get; set; }
            [Required]
            [DisplayName("檢查室名稱")]
            public string CalendarGroupName { get; set; }
            [Required]
            [DisplayName("病歷號")]
            public string MedicalNoteNo { get; set; }
            [Required]
            [DisplayName("檢查單號")]
            public string ExaRequestNo { get; set; }
            [Required]
            [DisplayName("檢查時間")]
            public DateTime Start { get; set; }
            public string ReservationSourceType { get; set; }
            public string SourceCode { get; set; }
            [Required]
            [DisplayName("主機病歷號")]
            public string DVC_CHRT { get; set; }
            [Required]
            [DisplayName("主機單號")]
            public string DVC_RQNO { get; set; }
            [Required]
            [DisplayName("主機排程日")]
            public string DVC_DATE { get; set; }
            [Required]
            [DisplayName("主機排程時間")]
            public string DVC_STTM { get; set; }
            [Required]
            [DisplayName("主機檢查碼1")]
            public string XRYSourceCode { get; set; }
            public DateTime PlanDate { get; set; }
        }
        public class ImportDBData
        {
            private ExcelWorksheet _sheet { get; set; }
            private int _rowIndex { get; set; }
            private int _colIndex { get; set; }
            private DataTable _dt { get; set; }
            private List<DBData> _dblist { get; set; }
            public void ImportData(DataTable dt, ExcelWorksheet sheet, int rowIndex, int colIndex, List<DBData> dblist)
            {
                _sheet = sheet;
                _rowIndex = rowIndex;
                _colIndex = colIndex;
                _dt = dt;
                _dblist = dblist;
                _sheet.Cells[_rowIndex - 1, _colIndex].Value = "返回目錄";
                _sheet.Cells[_rowIndex - 1, _colIndex].SetHyperlink(new Uri($"#'目錄'!A1", UriKind.Relative));

                //3.1塞columnName 到Row 
                for (int columnNameIndex = 0; columnNameIndex <= _dt.Columns.Count - 1; columnNameIndex++)
                {
                    MemberInfo property = typeof(DBData).GetProperty((_dt.Columns[columnNameIndex].ColumnName == null ? string.Empty : _dt.Columns[columnNameIndex].ColumnName));
                    var attribute = property.GetCustomAttributes(typeof(DisplayNameAttribute), true)
                                            .Cast<DisplayNameAttribute>().Single();
                    string columnName = attribute.DisplayName;
                    _sheet.Cells[_rowIndex, _colIndex++].Value = columnName;


                }
                _sheet.Cells[_rowIndex, 1, _rowIndex, _colIndex - 1]
                     .SetQuickStyle(Color.Black, Color.LightPink, ExcelHorizontalAlignment.Center);

                //將對應值放入
                foreach (var dbdata in _dblist)
                {
                    if (_sheet.ToString() == (dbdata.Start != DateTime.MinValue ? dbdata.Start.ToString("yyyy-MM-dd") : dbdata.PlanDate.ToString("yyyy-MM-dd")))
                    {
                        _rowIndex++;
                        _colIndex = 1;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.RESRoomCode;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.XRYRoomCode;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.CalendarGroupName;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.MedicalNoteNo;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.ExaRequestNo;
                        _sheet.Cells[_rowIndex, _colIndex].Value = dbdata.Start;
                        _sheet.Cells[_rowIndex, _colIndex].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        _sheet.Cells[_rowIndex, _colIndex++].Style.Numberformat.Format = "yyyy/MM/dd HH:mm:ss";
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.DVC_CHRT;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.DVC_RQNO;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.DVC_DATE;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.DVC_STTM;
                    }
                }

                //Autofit
                int startColumn = _sheet.Dimension.Start.Column;
                int endColumn = _sheet.Dimension.End.Column;
                for (int count = startColumn; count <= endColumn; count++)
                {
                    _sheet.Column(count).AutoFit();
                }


            }
            public void GenFirstSheet(ExcelPackage excel, List<DateTime> list)
            {
                int rowIndex = 1;
                int colIndex = 1;

                int maxCol = 0;

                ExcelWorksheet firstSheet = excel.Workbook.Worksheets.Add("目錄");

                firstSheet.Cells[rowIndex, colIndex++].Value = "";
                firstSheet.Cells[rowIndex, colIndex++].Value = "檢查時間";

                firstSheet.Cells[rowIndex, 1, rowIndex, colIndex - 1]
                    .SetQuickStyle(Color.Black, Color.LightPink, ExcelHorizontalAlignment.Center);

                maxCol = Math.Max(maxCol, colIndex - 1);

                foreach (DateTime info in list)
                {
                    rowIndex++;
                    colIndex = 1;

                    firstSheet.Cells[rowIndex, colIndex++].Value = rowIndex - 1;
                    firstSheet.Cells[rowIndex, colIndex++].Value = info.ToString("yyyy-MM-dd");
                    firstSheet.Cells[rowIndex, colIndex - 1].SetHyperlink(new Uri($"#'{(string.IsNullOrEmpty(info.ToString("yyyy-MM-dd")) ? info.ToString("yyyy-MM-dd") : info.ToString("yyyy-MM-dd"))}'!A1", UriKind.Relative));
                }

                for (int i = 1; i <= maxCol; i++)
                {
                    firstSheet.Column(i).AutoFit();
                }
            }

        }


    }
}
