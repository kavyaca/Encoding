using System;

namespace CSV.Models
{
    public class Constants
    {

        public  Student Student = new Student { StudentId = "200429439", FirstName = "Kavya", LastName = "Arora" };

        public class Locations
        {
            public readonly static string DesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            public readonly static string ExePath = Environment.CurrentDirectory;

            public readonly static string ContentFolder = $"{ExePath}\\..\\..\\..\\Content";
            public readonly static string DataFolder = $"{ContentFolder}\\Data";
            public readonly static string ImagesFolder = $"{ContentFolder}\\Images";

            public const string InfoFile = "info.csv";
            public const string ImageFile = "myimage.jpg";
        }

        public class FTP
        {
            public const string Username = @"bdat100119f\bdat1001";
            public const string Password = "bdat1001";
            public const string mydirectory = "200429439 Kavya Arora/info.csv";


            public const string BaseUrl = "ftp://bdat100119f%255Cbdat1001@waws-prod-dm1-127.ftp.azurewebsites.windows.net/BDAT1001-20914/";

            public const int OperationPauseTime = 10000;
        }
    }
}
