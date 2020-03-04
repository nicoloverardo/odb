using System.Diagnostics;
using System.Linq;

namespace OutlookDataBackup
{
    /// <summary> 
    /// Class that gets the network traffic from the performance counter. 
    /// Based on: http://pastebin.com/f371375d6 
    /// </summary> 
    public class NetworkTraffic
    {
        private PerformanceCounter bytesSentPerformanceCounter;
        private readonly int pid;
        private bool countersInitialized;

        public NetworkTraffic(int processID)
        {
            pid = processID;
            TryToInitializeCounters();
        }

        private void TryToInitializeCounters()
        {
            if (countersInitialized) return;

            var category = new PerformanceCounterCategory(".NET CLR Networking 4.0.0.0");
            var instanceNames = category.GetInstanceNames().Where(i => i.Contains($"p{pid}")).ToList();

            if (!instanceNames.Any()) return;

            bytesSentPerformanceCounter = new PerformanceCounter
            {
                CategoryName = ".NET CLR Networking 4.0.0.0",
                CounterName = "Bytes Sent",
                InstanceName = instanceNames.First(),
                ReadOnly = true
            };

            countersInitialized = true;
        }

        public float GetBytesSent()
        {
            try
            {
                TryToInitializeCounters();
                return bytesSentPerformanceCounter.RawValue;
            }
            catch
            {
                return 0;
            }
        }
    }
}
