using System.Diagnostics;
using System.Reflection;
using System.Threading;

namespace EarTrumpet.Interop.Helpers
{
    class SingleInstanceAppMutex 
    {
        private static Mutex s_mutex;

        public static bool TakeExclusivity()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var mutexName = $"Local\\{assembly.GetName().Name}-a5d1bbcd-8f69-41c7-be73-96d237374935";

            s_mutex = new Mutex(true, mutexName, out bool mutexCreated);
            if (!mutexCreated)
            {
                Trace.WriteLine("SingleInstanceAppMutex TakeExclusivity: false");
                s_mutex = null;
                return false;
            }
            return true;
        }

        public static void ReleaseExclusivity()
        {
            s_mutex?.ReleaseMutex();
            s_mutex?.Close();
            s_mutex = null;
        }
    }
}
