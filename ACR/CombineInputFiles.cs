﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;

namespace ACR
{
    class CombineInputFiles
    {
        //Name of file to combine
        string file_ams = @"Import\AMS.txt";
        string file_apac = @"Import\APAC.txt";
        string file_emea_bcn = @"Import\EMEA_BCN.txt";
        string file_emea_prg = @"Import\EMEA_PRG.txt";
        string file_GWFM_all_region = @"Import\ACR.txt";

        public void CreateFile()
        {
            if (File.Exists(file_ams))
            {
                ams();
            }

            if (File.Exists(file_apac))
            {
                apac();
            }

            if (File.Exists(file_emea_bcn))
            {
                emea_bcn();
            }

            if (File.Exists(file_emea_prg))
            {
                emea_prg();
            }
            
            if (File.Exists(file_GWFM_all_region))
            {
                GWFM_all();
            }

            string outputFile = @"Import\ACR.csv";
            File.Delete(outputFile);
            FileInfo CombinedFile_Output = new FileInfo(outputFile);
            File.WriteAllLines(outputFile, combined.Cast<string>());
        }

        //file array - will be output
        ArrayList combined = new ArrayList();

        private void ams()
        {
            string[] rawData = System.IO.File.ReadAllLines(file_ams);
            int count = 0;
            foreach (string rawRecord in rawData)
            {
                //string[] temp = rawRecord.Split('');
                string temp = rawRecord;
                temp = rawRecord.Replace("\"", "");
                combined.Add(temp);
                //count++;
            }
        }

        private void apac()
        {
            string[] rawData = System.IO.File.ReadAllLines(file_apac);
            int count = 0;
            foreach (string rawRecord in rawData)
            {
                //string[] temp = rawRecord.Split('');
                string temp = rawRecord;
                temp = rawRecord.Replace("\"", "");
                combined.Add(temp);
                //count++;
            }
        }

        private void emea_bcn()
        {
            string[] rawData = System.IO.File.ReadAllLines(file_emea_bcn);
            int count = 0;
            foreach (string rawRecord in rawData)
            {
                //string[] temp = rawRecord.Split('');
                string temp = rawRecord;
                temp = rawRecord.Replace("\"", "");
                combined.Add(temp);
                //count++;
            }
        }

        private void emea_prg()
        {
            string[] rawData = System.IO.File.ReadAllLines(file_emea_prg);
            int count = 0;
            foreach (string rawRecord in rawData)
            {
                //string[] temp = rawRecord.Split('');
                string temp = rawRecord;
                temp = rawRecord.Replace("\"", "");
                combined.Add(temp);
                //count++;
            }
        }

        private void GWFM_all()
        {
            string[] rawData = System.IO.File.ReadAllLines(file_GWFM_all_region);
            int count = 0;
            foreach (string rawRecord in rawData)
            {
                //string[] temp = rawRecord.Split('');
                string temp = rawRecord;
                temp = rawRecord.Replace("\"", "");
                combined.Add(temp);
                //count++;
            }
        }
    }
}
