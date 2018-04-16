using System.Runtime.InteropServices;

namespace Bar
{
    [ComVisible(true)]
    [Guid("8A29915D-2700-457C-A8DD-A682D22FF8A8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IEvents
    {
        [DispId(1)]
        void OnCounterChanged(int counter);
    }

    public delegate void OnCounterChangedDelegate(int counter);
}