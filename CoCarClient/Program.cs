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
        int Invoke([In, MarshalAs(UnmanagedType.U4)] int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U2)] short wFlags,
            System.Runtime.InteropServices.ComTypes.DISPPARAMS Params, dynamic /* */ pVarResult,
            [Out] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.U4)] int puArgErr);
    };
    class Program
    {        
        [STAThread]
        static void Main(string[] args)
        {
            dynamic kek=null;
            int a=0;
            Car myCar = new Car();
            Type type = typeof(Car);
            ICreateCar iCrCar = (ICreateCar)myCar;
            //Console.WriteLine("Напишите имя: ");
            //iCrCar.SetPetName(Console.ReadLine());
            //IStats iStCar = (IStats)myCar;
            //string st = "";
            //iStCar.GetPetName(ref st);
            //Console.WriteLine(st);
            IDispatch disp = (IDispatch)myCar;
            disp.GetTypeInfoCount(out a);
            Console.WriteLine( a.ToString());            

            int[] t = new int[1];
            Guid guid = new Guid();
            disp.GetIDsOfNames(ref guid, new string[] { "GetMaxSpeed" }, 1, 1, t);
            Console.WriteLine(t[0].ToString());

            
            System.Runtime.InteropServices.ComTypes.DISPPARAMS par = new System.Runtime.InteropServices.ComTypes.DISPPARAMS();
            int res = 1;
            System.Runtime.InteropServices.ComTypes.EXCEPINFO info = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
            //iDpCar.Invoke(t[0], guid, 1, 0, par, res, info, 0);
            int speed = 0;
            //type.InvokeMember("", BindingFlags.DeclaredOnly |
            //    BindingFlags.Public | BindingFlags.NonPublic |
            //    BindingFlags.Instance | BindingFlags.SetProperty, null, myCar, new object[] { a });
            //myCar.GetType().InvokeMember("Invoke", BindingFlags.DeclaredOnly |
            //    BindingFlags.Public | BindingFlags.NonPublic |
            //    BindingFlags.Instance | BindingFlags.SetProperty, null, myCar, new object[] { t[0], guid, 1, 0, par, kek, info, 0 });
            Console.WriteLine(speed.ToString());
            Console.ReadKey();
        }
    }
}
