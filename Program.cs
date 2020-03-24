using CSV.Models;
using CSV.Models.Utilities;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;

namespace CSV
{
    class Program
    {
        private static int encryptionDepth = 20;
        static void Main(string[] args)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            Console.WriteLine(FTP.DownloadFile(Constants.FTP.BaseUrl+Constants.FTP.mydirectory,desktopPath));

            string filepath = desktopPath+"/info.csv";
            string fileContents;

            using (StreamReader stream = new StreamReader(filepath))
            {
                fileContents = stream.ReadToEnd();
            }

            List<String> entries = new List<string>();
            entries = fileContents.Split("\r\n", StringSplitOptions.RemoveEmptyEntries).ToList();
            String[] Data = entries[1].Split(',');
            

            String encryptionKey = Data[1];

            int[] key = new int[5];
            int i = 0;
            foreach (char c in encryptionKey)
            {
                key[i] = System.Convert.ToInt32(c);
                i = i + 1;
            }
            //Encryption Decryption
            string myString = String.Join(",", key.Select(x => x.ToString()));
            Console.WriteLine(myString);


       
            string[] nameDeepEncryptWithCipher = new string[5];
            i = 0;

            foreach (string letter in Data)
            {
                Encrypter encrypter = new Encrypter(letter, key, encryptionDepth);
                //Deep Encrytion
                nameDeepEncryptWithCipher[i] = Encrypter.DeepEncryptWithCipher(letter, key, encryptionDepth);
                Console.WriteLine($"\nDeep Decrypted {encryptionDepth} times using the cipher {Data[i]} {nameDeepEncryptWithCipher[i]}");
                i = i + 1;
            }

            string finalstring = String.Join(",", nameDeepEncryptWithCipher);
            var csv = new StringBuilder();

            csv.AppendLine(finalstring);

            File.WriteAllText(desktopPath, csv.ToString());
        }

    }
}
