using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            _timer = new Timer(1000) { AutoReset = true };
            _timer.Elapsed += TimerElapsed;
        }

        private void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            ProcessFiles oProcessFiles = new ProcessFiles();
            oProcessFiles.ProcessingFile();
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
