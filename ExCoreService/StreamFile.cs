using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace ExCoreService
{
    public class StreamFile
    {
        private readonly Timer _timer;

        public StreamFile()
        {
            _timer = new Timer(AppSettings.TimerPollingInterval) { AutoReset = true };
            _timer.Elapsed += TimerElapsed;
        }

        private void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                _timer.Stop();
                ProcessFiles oProcessFiles = new ProcessFiles();
                oProcessFiles.ProcessingFile("MfgOrder", AppSettings.OrderSourceFolder, AppSettings.OrderCompletedFolder, AppSettings.OrderErrorFolder);
                oProcessFiles.ProcessingFile("OrderBOM", AppSettings.OrderBOMSourceFolder, AppSettings.OrderBOMCompletedFolder, AppSettings.OrderBOMErrorFolder);
            }
            catch(Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                EventLogUtil.LogErrorEvent(ex.Source, ex);
            }
            finally
            {
                _timer.Start();
            }
        }

        public void Start()
        {
            _timer.Start();
        }

        public void Stop()
        {
            _timer.Stop();
        }
    }
}
