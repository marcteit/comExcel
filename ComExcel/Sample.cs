using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace ComExcel
{
    [Serializable, ClassInterface(ClassInterfaceType.AutoDual), ComVisible(true)]
    public class Sample
    {
        private Thread _mainThread;
        public int Counter { get; set; }

        public void Initialize()
        {
            _mainThread = Thread.CurrentThread;
            Task.Factory.StartNew(() =>
            {
                while (true)
                {
                    Counter++;
                    Thread.Sleep(1000);
                }
            });
        }

        public void Wait(int expectedCounter)
        {
            while (Counter <= expectedCounter)
            {
                Thread.Sleep(100);
            }
        }
    }

    static class UnmanagedExports
    {
        //[DllExport]
        [return: MarshalAs(UnmanagedType.IDispatch)]
        static Object CreateTestClass()
        {
            return new TestClass();
        }
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public class TestClass
    {

        public int Counter { get; set; }

        public void Initialize()
        {
            Task.Factory.StartNew(() =>
            {
                while (true)
                {
                    Counter++;
                    Thread.Sleep(1000);
                }
            });
        }

        public void Wait(int expectedCounter)
        {
            while (Counter <= expectedCounter)
            {
                Thread.Sleep(100);
            }
        }
    }
}
