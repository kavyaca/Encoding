using System;

namespace CSV.Models
{
    public class Constants
    {

        public static readonly Student Student = new Student { StudentId = "200429439", FirstName = "Kavya", LastName = "Arora" };

        public class Locations
        {

            public readonly static string api_url = "https://jsonplaceholder.typicode.com/users";
            public readonly static string DesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            public readonly static string ExePath = Environment.CurrentDirectory;
            public readonly static string ContentFolder = $"{ExePath}//..//..//..//Content";
            public readonly static string DataFolder = $"{ContentFolder}//Data";
            public readonly static string ImagesFolder = $"{ContentFolder}//Images";
            public readonly static string FilesFolder = $"{ContentFolder}//Files";
            public const string InfoFile = "info.csv";
            public const string ImageFile = "myimage.jpg";
            public readonly static string MyID = "200429439";
            public readonly static string StudentCSVFile = $"{DataFolder}//students.csv";
            public readonly static string StudentJSONFile = $"{DataFolder}//students.json";
            public readonly static string StudentXMLFile = $"{DataFolder}//students.xml";
            public readonly static string StudentDocFile = $"{DataFolder}//students.docx";
            public readonly static string StudentDocFileFromApi = $"{DataFolder}//studentsapi.docx";
            public readonly static string StudentExcelFile = $"{DataFolder}//students.xlsx";
            public readonly static string StudentPPTFile = $"{DataFolder}//students.pptx";


        }

        public class FTP
        {
            public const string Username = @"bdat100119f\bdat1001";
            public const string Password = "bdat1001";

            public const string BaseUrl = "ftp://waws-prod-dm1-127.ftp.azurewebsites.windows.net/bdat1001-20914";
            public const string MyDirectory = "/200429439 Kavya Arora";
            public const string CSVUploadLocation = BaseUrl + MyDirectory + "/students.csv";
            public const string JSONUploadLocation = BaseUrl + MyDirectory + "/students.json";
            public const string XMLUploadLocation = BaseUrl + MyDirectory + "/students.xml";

            public const int OperationPauseTime = 10000;

        }
    }
}
