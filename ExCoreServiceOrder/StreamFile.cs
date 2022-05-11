using OpcenterWikLibrary;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Timers;


namespace ExCoreServiceOrder
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
                ProcessFiles processFiles = new ProcessFiles();
                processFiles.QueuingFiles(AppSettings.SourceFolder, AppSettings.QueueFolder, AppSettings.CompletedFolder, AppSettings.ErrorFolder, AppSettings.ReferenceFolder);
            }
            catch(Exception ex)
            {
                ex.Source = AppSettings.AssemblyName == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
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
