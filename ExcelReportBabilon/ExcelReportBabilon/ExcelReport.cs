using System;
using System.Data;
using System.Data.SqlClient;
using System.Net.Mail;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;
using log4net;

namespace ExcelReportBabilon
{
    class ExcelReport
    {
        private static readonly string _connectString = @"Data Source=(local);database=;Integrated Security=True";
        private static readonly ILog _log = LogManager.GetLogger(typeof(ExcelReport));

        public static void Start(string dateBegin, string dateEnd)
        {
            Excel.Workbook ExcelWB = null;
            try
            {
                _log.Info("Start reporting:...");
                Excel.Application ExcelApp = new Excel.Application();

                ExcelWB = ExcelApp.Workbooks.Open(@"E:\Babilon Excel report\отчёт_babilon.xlsx");

                Excel._Worksheet ExcelWS = (Excel._Worksheet)(ExcelWB.ActiveSheet);
                _log.Info("Get payments:...");
                DataTable dt = GetData("GetPaymentsBabilon", new SqlParameter[]
                {
                    new SqlParameter("@dateBegin", dateBegin),
                    new SqlParameter("@dateEnd", dateEnd)
                });

                int iRow = 4;
                _log.Info("filling excel file:...");
                foreach (DataRow row in dt.Rows)
                {
                    ExcelWS.Cells[iRow, 2] = row["RegDateTime"].ToString();
                    ExcelWS.Cells[iRow, 3] = row["PaymentID"].ToString();

                    ExcelWS.Cells[iRow, 4] = row["Number"].ToString();
                    ExcelWS.Cells[iRow, 5] = row["PaySum"].ToString();
                    ExcelWS.Cells[iRow, 6] = row["Status"].ToString();
                    Excel.Range range = ExcelWS.Range[ExcelWS.Cells[iRow, 2], ExcelWS.Cells[iRow, 6]];
                    range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                    ++iRow;
                }
                _log.Info("Saving excel file:...");
                ExcelWB.SaveAs($"E:\\Babilon Excel report\\отчёт_babilon_{DateTime.Now.AddDays(-1.0).ToString("yyyy-MM-dd")}.xlsx");
                ExcelWB.Close();
                _log.Info("Sending mail:...");
                SendMail("", "");

            }
            catch (Exception ex)
            {
                ExcelWB?.Close();
                _log.Error(ex);
            }
        }

        private static void SendMail(string from, string to)
        {
            MailMessage message = new MailMessage(from, to);
            message.Subject = "Отчёт";
            message.Body = "<h2></h2>";

            message.Attachments.Add(new Attachment($"E:\\Babilon Excel report\\отчёт_babilon_{DateTime.Now.AddDays(-1.0).ToString("yyyy-MM-dd")}.xlsx"));

            message.IsBodyHtml = true;
            SmtpClient smtp = new SmtpClient("", 25);
            smtp.Credentials = new NetworkCredential("", "");
            smtp.EnableSsl = false;
            smtp.Send(message);
        }

        private static DataTable GetData(string SPName, SqlParameter[] SQLParam)
        {
            using (SqlConnection ConnectToDB = new SqlConnection(_connectString))
            {
                ConnectToDB.Open();
                SqlCommand SQLCommand = new SqlCommand(SPName, ConnectToDB);
                SQLCommand.CommandType = CommandType.StoredProcedure;
                if (SQLParam != null)
                {
                    for (int i = 0; i < SQLParam.Length; i++)
                    {
                        SQLCommand.Parameters.Add(SQLParam[i]);
                    }
                }
                using (SqlDataAdapter SQLDataAdapter = new SqlDataAdapter(SQLCommand))
                {
                    DataSet dataSet = new DataSet();
                    SQLDataAdapter.Fill(dataSet);
                    return dataSet.Tables[0];
                }
            }
        }

    }
}
