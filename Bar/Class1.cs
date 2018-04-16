using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.EnterpriseServices;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Bar
{
    [Serializable, ClassInterface(ClassInterfaceType.AutoDual), ComVisible(true)]
    [ComClass("B1203BCC-F1C6-42C7-8EC1-0E719865F74C", "11203BCC-F1C6-42C7-8EC1-0E719865F74C", "21203BCC-F1C6-42C7-8EC1-0E719865F74C")]
    [ComSourceInterfaces(typeof(IEvents))]
    public class TestClassBase : ServicedComponent
    {
        public int Counter { get; set; }
        public virtual void Initialize()
        {

        }
        public virtual void Wait(int expectedCounter)
        {

        }
        public event OnCounterChangedDelegate OnCounterChanged;
        protected void FireOnCounterChanged()
        {
            if (OnCounterChanged != null)
            {
                OnCounterChanged(Counter);
            }
        }
    }
}
