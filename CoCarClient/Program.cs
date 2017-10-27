using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CoCarClient
{
    

    class Program
    {
        [ComVisible(true)]
        [ComImport, Guid("ECD55377-DFEF-4C48-BAD8-BB9D00C53295")]
        public class Car
        {
        }
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("60C8F2BC-B1A7-483C-BDB2-A42C49E0F932")]
        public interface ICreateCar
        {
            void SetPetName(string petName);
            void SetMaxSpeed(int maxSp);
        }
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("488084B6-A989-4196-9488-C4D1CAE40EDC")]
        public interface IStats
        {
            void DisplayStats();
            void GetPetName(ref string petName);
        }
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("140BB72A-172A-4133-9B3E-D7FC9E6CFD87")]
        public interface IEngine
        {
            void SpeedUp();
            void GetMaxSpeed(ref int curSpeed);
            void GetCurSpeed(ref int maxSpeed);
        }

        [STAThread]
        static void Main(string[] args)
        {
            Car myCar = new Car();
            ICreateCar iCrCar = (ICreateCar)myCar;
            Console.WriteLine("Напишите имя: ");
            iCrCar.SetPetName(Console.ReadLine());

        }
    }
}
