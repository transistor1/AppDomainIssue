using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Com_2
{

    [Guid("A9C880C6-6C97-4B02-8970-40483209ADA6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [ComVisible(true)]
    public interface IComClass2
    {
        void SetAppDomainData(string data);
        string GetAppDomainData();
        int GetAppDomainHash();
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("0351EFC7-137A-42A4-A62E-6FD402C2484F")]
    [ProgId("Com2.ComClass2")]
    public class ComClass2 : IComClass2
    {
        public void SetAppDomainData(string data)
        {
            AppDomain.CurrentDomain.SetData("CurrentDomainCustomData", data);
        }

        public string GetAppDomainData()
        {
            return (string)AppDomain.CurrentDomain.GetData("CurrentDomainCustomData");
        }

        public int GetAppDomainHash()
        {
            return AppDomain.CurrentDomain.GetHashCode();
        }
    }
}