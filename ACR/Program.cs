using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACR
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Welcome! Creating 100 Day Report...");
            run();
        }

        private static void run()
        {
            CombineInputFiles combine_files = new CombineInputFiles();
            combine_files.CreateFile();
            string file = @"Import\ACR.csv";
            if (File.Exists(file))
            {
                Console.WriteLine("\nFound the import file");

                int length = File.ReadLines(file).Count();
                ProcessFile input = new ProcessFile(length);
                input.RawRead();
            }
            else
            {
                Console.WriteLine("\nError! Import file NOT found at");
                Console.ReadKey();
            }
        }
        public void reRun()
        {
            run();
        }
    }
}
