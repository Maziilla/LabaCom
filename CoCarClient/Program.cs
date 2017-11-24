using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;

namespace CoCarClient
{

    [ComVisible(true)]
    [ComImport, Guid("ECD55377-DFEF-4C48-BAD8-BB9D00C53295")]
    public class Car
    {
    };
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("60C8F2BC-B1A7-483C-BDB2-A42C49E0F932")]
    public interface ICreateCar
    {
        void SetPetName(string petName);
        void SetMaxSpeed(int maxSp);
    };
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("488084B6-A989-4196-9488-C4D1CAE40EDC")]
    public interface IStats
    {
        void DisplayStats();
        void GetPetName(ref string petName);
    };
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("140BB72A-172A-4133-9B3E-D7FC9E6CFD87")]
    public interface IEngine
    {
        void SpeedUp();
        void GetMaxSpeed(ref int curSpeed);
        void GetCurSpeed(ref int maxSpeed);
    };
    [ComVisible(true)]
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("97747077-659A-43ED-BB74-9C125151BE8D")]
    public interface IDispatch
    {
        int GetTypeInfoCount(out int Count);
        [return: MarshalAs(UnmanagedType.Interface)]
        int GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] UInt32 iTInfo, [In, MarshalAs(UnmanagedType.U4)] UInt32 lcid, ref ITypeInfo typeInfo);
        int GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames,
            [In, MarshalAs(UnmanagedType.U4)] int cNames, [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        int Invoke([In, MarshalAs(UnmanagedType.U4)] int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, System.Runtime.InteropServices.ComTypes.INVOKEKIND wFlags,
            ref System.Runtime.InteropServices.ComTypes.DISPPARAMS Params, out object pVarResult,
            IntPtr pExcepInfo, IntPtr puArgErr);
    };
    class Program
    {        
        [STAThread]
        static void Main(string[] args)
        {
            object Res = 1;
            int a=0;
            Car myCar = new Car();
            //Type type = typeof(Car);
            ICreateCar iCrCar = (ICreateCar)myCar;
            IStats iStCar = (IStats)myCar;
            //Console.WriteLine("Напишите имя: ");
            //iCrCar.SetPetName(Console.ReadLine());
            //IStats iStCar = (IStats)myCar;
            //string st = "";
            //iStCar.GetPetName(ref st);
            //Console.WriteLine(st);
            iCrCar.SetPetName("Lolipop");
            iCrCar.SetMaxSpeed(67);
            IDispatch disp = (IDispatch)myCar;
            iStCar.DisplayStats();
            disp.GetTypeInfoCount(out a);
            
            Console.WriteLine( a.ToString());            

            int[] t = new int[1];
            Guid guid = new Guid();
            disp.GetIDsOfNames(ref guid, new string[] { "GetMaxSpeed" }, 1, 1, t);
            Console.WriteLine("Max speed "+t[0].ToString());

            var arg = 51;
            var pVariant = Marshal.AllocCoTaskMem(16);
            Marshal.GetNativeVariantForObject(arg, pVariant);
            System.Runtime.InteropServices.ComTypes.DISPPARAMS par = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
            {
                cArgs = 1,
                cNamedArgs = 0,
                rgdispidNamedArgs = IntPtr.Zero,
                rgvarg = pVariant
            };
            System.Runtime.InteropServices.ComTypes.INVOKEKIND flags = new System.Runtime.InteropServices.ComTypes.INVOKEKIND();
            IntPtr info = new IntPtr();
            IntPtr puArgErr = new IntPtr();
            disp.Invoke(4, guid, 1, flags, par, out Res, info, puArgErr);
            disp.Invoke(1, guid, 1, flags, par, out Res, info, puArgErr);
            Console.WriteLine("Res = " + Res.ToString());

            var arg_6 = "";
            Marshal.GetNativeVariantForObject(arg_6, pVariant);
            System.Runtime.InteropServices.ComTypes.DISPPARAMS par_6 = new System.Runtime.InteropServices.ComTypes.DISPPARAMS()
            {
                cArgs = 1,
                cNamedArgs = 0,
                rgdispidNamedArgs = IntPtr.Zero,
                rgvarg = pVariant
            };
            System.Runtime.InteropServices.ComTypes.INVOKEKIND flags_6 = new System.Runtime.InteropServices.ComTypes.INVOKEKIND();
            disp.Invoke(6, guid, 1, flags_6, par_6, out Res, info, puArgErr);


            Console.WriteLine("Res = "+Res.ToString());
            Console.ReadKey();
        }
    }
}
