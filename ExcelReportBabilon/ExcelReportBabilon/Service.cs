using log4net;
using log4net.Config;
using System;
using System.Globalization;
using System.IO;
using System.ServiceProcess;
using System.Threading;

namespace ExcelReportBabilon
{
    public partial class Service : ServiceBase
    {
        Thread _thread = null;
        static readonly ILog _log = LogManager.GetLogger(typeof(Service));

        public Service()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                _thread = new Thread(() => CheckTime());
                _thread.Start();
            }
            catch (Exception ex)
            {
                _log.Error(ex);
            }
        }

        void CheckTime()
        {
            XmlConfigurator.Configure(new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log4net.config")));
            bool status = false;
            _log.Info("Start ...");
            while (true)
            {
                DateTime.TryParseExact("00", "HH", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime t);

                if (t.ToString("HH") == DateTime.Now.ToString("HH"))
                {
                    if (status == false)
                    {
                        string dateBegin, dateEnd;
                        dateBegin = DateTime.Now.AddDays(-1.0).ToString("yyyy-MM-dd") + " 00:00:00";
                        dateEnd = DateTime.Now.ToString("yyyy-MM-dd") + " 00:00:00";
                        try
                        {
                            ExcelReport.Start(dateBegin, dateEnd);
                        }
                        catch(ThreadAbortException ex)
                        {
                            _log.Error(ex);
                            break;
                        }
                        catch(Exception ex)
                        {
                            _log.Error(ex);
                        }

                        status = true;
                    }
                }
                else
                {
                    status = false;
                }
                Thread.Sleep(100000);
            }
        }

        protected override void OnStop()
        {
            try
            {
                _thread?.Abort();
                _thread?.Join();
            }
            catch(Exception ex)
            {
                _log.Error(ex);
            }
        }
    }
}
