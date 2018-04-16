using Microsoft.VisualBasic;
using RGiesecke.DllExport;
using System;
using System.EnterpriseServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace Foo
{
    public static class UnmanagedExports
    {
        [DllExport]
        [return: MarshalAs(UnmanagedType.IDispatch)]
        static Object CreateTestClass()
        {
            return new TestClass();
        }
    }

    public delegate void OnCounterChangedDelegate(int counter);

    [ComVisible(true)]
    [Guid("235C2869-3BE6-410A-A8C9-1D3042339C9C")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ITestClassEvents
    {
        [DispId(1)]
        void OnCounterChanged(int counter);
    }


    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    [ComClass("62026A8D-0D0A-49ED-9B98-C318E409361F", "72026A8D-0D0A-49ED-9B98-C318E409361F", "82026A8D-0D0A-49ED-9B98-C318E409361F")]
    [ComSourceInterfaces(typeof(ITestClassEvents))]
    public class TestClass //: ServicedComponent
    {
        public event OnCounterChangedDelegate OnCounterChanged;
        public int Counter { get; private set; }
        public void Initialize()
        {
            Task.Factory.StartNew(() =>
            {
                while (true)
                {
                    Counter++;
                    if (OnCounterChanged != null)
                    {
                        OnCounterChanged(Counter);
                    }
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
