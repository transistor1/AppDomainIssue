﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ComClasses
{
    internal class GenericClassFactory<comType, interfaceType> : IClassFactory
    {
        public int CreateInstance(IntPtr pUnkOuter, ref Guid riid,
            out IntPtr ppvObject)
        {
            ppvObject = IntPtr.Zero;

            if (pUnkOuter != IntPtr.Zero)
            {
                // The pUnkOuter parameter was non-NULL and the object does 
                // not support aggregation.
                Marshal.ThrowExceptionForHR(COMNative.CLASS_E_NOAGGREGATION);
            }

            FieldInfo clsIdField = typeof(comType).GetField("ClassId", BindingFlags.Static |
                BindingFlags.NonPublic | BindingFlags.Public );
            string clsId = (string)clsIdField?.GetValue(null);

            if (clsId == null)
                throw new NotImplementedException("The ClassId field is not implemented in class " + typeof(comType).Name + ". Please make sure this class declares a ClassId string, containing the class's GUID." );

            if (riid == new Guid(clsId) ||
                riid == new Guid(COMNative.IID_IDispatch) ||
                riid == new Guid(COMNative.IID_IUnknown))
            {
                //Get the default constructor
                ConstructorInfo comConstructor = typeof(comType).GetConstructor(Type.EmptyTypes);
                comType comInstance = (comType)comConstructor.Invoke(null);

                // Create the instance of the .NET object
                ppvObject = Marshal.GetComInterfaceForObject(
                    comInstance, typeof(interfaceType));
            }
            else
            {
                // The object that ppvObject points to does not support the 
                // interface identified by riid.
                Marshal.ThrowExceptionForHR(COMNative.E_NOINTERFACE);
            }

            return 0;   // S_OK
        }

        public int LockServer(bool fLock)
        {
            return 0;   // S_OK
        }
    }
}
 