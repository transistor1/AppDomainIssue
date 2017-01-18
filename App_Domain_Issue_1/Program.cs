using Com_1;
using ComClasses;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

class Program
{
    /// <summary>
    /// The main entry point for the application.
    /// </summary>
    static void Main(string[] args)
    {
        // Run the out-of-process COM server
        OutOfProcServer.Instance.Run(ComClass1.ClassId, typeof(ComClass1ClassFactory));
    }
}

