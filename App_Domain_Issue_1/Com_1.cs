using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Com_1
{

    [Guid(ComClass1.InterfaceId)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [ComVisible(true)]
    public interface IComClass1
    {
        void SetAppDomainData(string data);
        string GetAppDomainData();
        int GetAppDomainHash();
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid(ComClass1.ClassId)]
    [Description("Sample COM Class 1")]
    [ProgId("Com1.ComClass1")]
    public class ComClass1 : IComClass1
    {
        internal const string ClassId = "3CA12D49-CFE5-45A3-B114-22DF2D7A0CAB";
        internal const string InterfaceId = "F35D5D5D-4A3C-4042-AC35-CE0C57AF8383";

        static ComClass1()
        {
        }

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
