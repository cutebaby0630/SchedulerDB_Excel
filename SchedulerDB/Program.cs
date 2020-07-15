using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.ComponentModel.Design;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography.Pkcs;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Style;
using ServiceStack.Text;
using SqlServerHelper.Core;


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

                             select * from #RESTTReservation 
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
            int rowCount = (dt == null) ? 0 : dt.Rows.Count;

            Console.WriteLine(rowCount);
            //Step 1.1.將資料放入List
            List<Data> migrationTableInfoList = sqlHelper.QueryAsync<Data>(sql).Result?.ToList();
            //Step 1.2 將Date Distinct  遞增 order by 遞減OrderByDescending
            var datetime = migrationTableInfoList.Select(p => p.Start != DateTime.MinValue ? p.Start.Date : p.PlanDate.Date)
                                                 .OrderBy(p => p.Date)
                                                 .Distinct()
                                                 .ToList();




            //Step 1.3.Group by start 塞入新的list
            /*  var result = from sqllist in migrationTableInfoList
                           join date in datetime
                           on (sqllist.Start != DateTime.MinValue) ? sqllist.Start : sqllist.PlanDate equals date into map
                           from allresult in map
                           select new { sqllist.Start, sqllist.PlanDate, date = allresult.ToString()};

              foreach (var x in result)
              {
                  Console.WriteLine(x);
              }*/
            //Step 2.建立 各日期Sheet
            var excelname = "Scheduler" + DateTime.Now.ToString("yyyyMMddhhmm") + ".xlsx";
            var excel = new FileInfo(excelname);
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var finish = new ExcelPackage(excel))
            {
                for (int num = 0; num <= datetime.Count - 1; num++)
                {
                    finish.Workbook.Worksheets.Add(datetime[num].ToString("yyyy-MM-dd"));
                }
                Byte[] bin = finish.GetAsByteArray();
                File.WriteAllBytes(@"D:\微軟MCS\SchedulerDB_Excel\" + excelname, bin);

            }
            //Step 3.將對應的List 丟到各Sheet中
            FileInfo excel_new = new FileInfo(@"D:\微軟MCS\SchedulerDB_Excel\" + excelname);
            using (ExcelPackage package = new ExcelPackage(excel_new))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets[0];
                int rowIndex = 1;
                int colIndex = 1;

                //3.1塞資料到某一格
                sheet.Cells[rowIndex, colIndex++].Value = "新檢查室";
                sheet.Cells[rowIndex, colIndex++].Value = "舊檢查室";
                sheet.Cells[rowIndex, colIndex++].Value = "檢查室名稱";
                sheet.Cells[rowIndex, colIndex++].Value = "病歷號";
                sheet.Cells[rowIndex, colIndex++].Value = "檢查單號";
                sheet.Cells[rowIndex, colIndex++].Value = "檢查時間";
                sheet.Cells[rowIndex, colIndex++].Value = "主機病歷號";
                sheet.Cells[rowIndex, colIndex++].Value = "主機單號";
                sheet.Cells[rowIndex, colIndex++].Value = "主機排程日";
                sheet.Cells[rowIndex, colIndex++].Value = "主機排程時間";
                sheet.Cells[rowIndex, colIndex++].Value = "主機檢查碼1";
                sheet.Cells[rowIndex, colIndex++].Value = "主機檢查碼2";
                sheet.Cells[rowIndex, colIndex++].Value = "主機檢查碼3";
                sheet.Cells[rowIndex, 1, rowIndex, colIndex - 1]
                     .SetQuickStyle(Color.Black, Color.LightPink, ExcelHorizontalAlignment.Center);

                //將對應值放入
                foreach (var v in migrationTableInfoList)
                {
                    if (sheet.ToString() == (v.Start != DateTime.MinValue ? v.Start.ToString("yyyy-MM-dd") : v.PlanDate.ToString("yyyy-MM-dd")))
                    {
                        rowIndex++;
                        colIndex = 1;
                        sheet.Cells[rowIndex, colIndex++].Value = v.RESRoomCode;
                        sheet.Cells[rowIndex, colIndex++].Value = v.XRYRoomCode;
                        sheet.Cells[rowIndex, colIndex++].Value = v.CalendarGroupName;
                        sheet.Cells[rowIndex, colIndex++].Value = v.MedicalNoteNo;
                        sheet.Cells[rowIndex, colIndex++].Value = v.ExaRequestNo;
                        sheet.Cells[rowIndex, colIndex].Value = v.Start;
                        sheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        sheet.Cells[rowIndex, colIndex++].Style.Numberformat.Format = "yyyy/MM/dd HH:mm:ss";
                        sheet.Cells[rowIndex, colIndex++].Value = v.DVC_CHRT;
                        sheet.Cells[rowIndex, colIndex++].Value = v.DVC_RQNO;
                        sheet.Cells[rowIndex, colIndex++].Value = v.DVC_DATE;
                        sheet.Cells[rowIndex, colIndex++].Value = v.DVC_STTM;
                    }
                }
                //2020/6/15 上午 12:00:00
                Console.WriteLine(DateTime.Parse(sheet.ToString()));
                //2020 - 06 - 15
                Console.WriteLine(sheet);

                //Autofit
                int startColumn = sheet.Dimension.Start.Column;
                int endColumn = sheet.Dimension.End.Column;
                for (int count = startColumn; count <= endColumn; count++)
                {
                    sheet.Column(count).AutoFit();
                }
                Byte[] bin = package.GetAsByteArray();
                File.WriteAllBytes(@"D:\微軟MCS\SchedulerDB_Excel\" + excelname, bin);
            }
            //Step 4.Export EXCEL
        }

        public class Data
        {
            public string RESRoomCode { get; set; }
            public string XRYRoomCode { get; set; }
            public string CalendarGroupName { get; set; }
            public string MedicalNoteNo { get; set; }
            public string ExaRequestNo { get; set; }
            public DateTime Start { get; set; }
            public string ReservationSourceType { get; set; }
            public string SourceCode { get; set; }
            public string DVC_CHRT { get; set; }
            public string DVC_RQNO { get; set; }
            public string DVC_DATE { get; set; }
            public string DVC_STTM { get; set; }
            public string XRYSourceCode { get; set; }
            public DateTime PlanDate { get; set; }
        }

    }
}
