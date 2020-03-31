using CSV.Models;
using CSV.Models.Utilities;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

using A = DocumentFormat.OpenXml.Drawing;

using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;



namespace CSV
{
    class Program
    {
        static void Main(string[] args)
        {



            List<string> directories = new List<string>();
            directories = FTP.GetDirectory(Constants.FTP.BaseUrl);
            //List<string> directories = FTP.GetDirectory(Constants.FTP.BaseUrl);
            List<Student> students = new List<Student>();
            int processed = 0;



            foreach (var directory in directories)
            {
                ++processed;
                if (processed == 10) break;


                Console.WriteLine("Directory: " + directory);
                Student student = new Student() { AbsoluteUrl = Constants.FTP.BaseUrl };
                student.FromDirectory(directory);


                // Console.WriteLine(student);
                string infoFilePath = student.FullPathUrl + "/" + Constants.Locations.InfoFile;

                bool fileExists = FTP.FileExists(infoFilePath);
                if (fileExists == true)
                {
                    //QUESTION 2
                    Console.WriteLine("Found info file: " + infoFilePath);
                    string firstname = student.FirstName;
                    //Console.WriteLine("Student name:" + firstname);
                    if (student.StudentId == Constants.Student.StudentId)
                    {
                        //student.FirstName
                        student.MyRecord = true;


                    }
                    else
                    {
                        student.MyRecord = false;
                    }


                    var infoBytes = FTP.DownloadFileBytes(infoFilePath);

                    //CSV data
                    string csv = Encoding.Default.GetString(infoBytes);



                    string[] csv_content = csv.Split("\r\n", StringSplitOptions.RemoveEmptyEntries);
                    if (csv_content.Length != 2)
                    {
                        Console.WriteLine("Error in CSV format.");

                    }
                    else
                    {

                        student.FromCSV(csv_content[1]);
                        //Console.WriteLine("Extracting data from "+student.FirstName+ " directory: \n" + csv_content[1]);



                    }

                }
                else
                {
                    //QUESTION 2
                    Console.WriteLine("Could not find info file:");
                    try
                    {
                        if (student.StudentId == "200429439")
                        {

                            //student.FirstName
                            student.MyRecord = true;


                        }
                        else
                        {
                            student.MyRecord = false;
                        }
                    }
                    catch (Exception ex)
                    {

                    }
                }

                //Console.WriteLine("Info File Path::: " + infoFilePath);

                string imageFilePath = student.FullPathUrl + "/" + Constants.Locations.ImageFile;

                bool imageFileExists = FTP.FileExists(imageFilePath);

                if (imageFileExists == true)
                {

                    Console.WriteLine("Found image file: " + imageFilePath);

                    if ((student.ImageData) == null || student.ImageData.Length < 2)
                    {
                        Console.WriteLine("No Imagedata:");



                        try
                        {
                            var ImageBytes = FTP.DownloadFileBytes(imageFilePath);
                            string base64String = Convert.ToBase64String(ImageBytes);
                            student.ImageData = base64String;



                        }
                        catch (Exception ex)
                        {
                            student.ImageData = "No data Found";
                        }
                    }



                    //Console.WriteLine("Image File Path::: " + imageFilePath);


                }
                else
                {
                    //QUESTION 2
                    Console.WriteLine("Could not find image file:");

                    try
                    {
                        student.ImageData = "No data Found";
                    }
                    catch (Exception ex)
                    {

                    }

                }


                students.Add(student);

                //}

            }
            //Console.WriteLine("Directory count: " + directories.Count);

            Console.WriteLine(Constants.Locations.StudentDocFile);
            Console.WriteLine(Constants.Locations.StudentExcelFile);

            SpreadSheetClass.CreateSpreadsheetWorkbook(Constants.Locations.StudentExcelFile, students);
            WordDocumentClass.CreateWordprocessingDocument(Constants.Locations.StudentDocFile, students);

           



            
            using (StreamWriter fs = new StreamWriter(Constants.Locations.StudentCSVFile))
            {
                //int Age = 10;
                fs.WriteLine((nameof(Student.StudentId)) + ',' + (nameof(Student.FirstName)) + ',' + (nameof(Student.LastName)) + ',' + (nameof(Student.Age)) + ',' + (nameof(Student.DateOfBirth)) + ',' + (nameof(Student.MyRecord)) + ',' + (nameof(Student.ImageData)));
                foreach (var student in students)
                {
                    fs.WriteLine(student.ToCSV());
                    Console.WriteLine("To CSV :: " + student.ToCSV());
                   Console.WriteLine("To String :: " + student.ToString());


                }
            }





            String jsonconvert = ConvertCsvFileToJsonObject(Constants.Locations.StudentCSVFile);

            File.WriteAllText(Constants.Locations.StudentJSONFile, jsonconvert);





            //TEST CSV TO XML
            //START



            var lines = File.ReadAllLines(Constants.Locations.StudentCSVFile);
            string[] headers = lines[0].Split(',').Select(x => x.Trim('\"')).ToArray();

            var xml = new XElement("root",
               lines.Where((line, index) => index > 0).Select(line => new XElement("row",
                  line.Split(',').Select((column, index) => new XElement(headers[index], column)))));

            xml.Save(Constants.Locations.StudentXMLFile);
            //END


            AggregateFunctions(students);



            FTP.UploadFile(Constants.Locations.StudentCSVFile, Constants.FTP.CSVUploadLocation);
            FTP.UploadFile(Constants.Locations.StudentXMLFile, Constants.FTP.XMLUploadLocation);
            FTP.UploadFile(Constants.Locations.StudentJSONFile, Constants.FTP.JSONUploadLocation);

            
            return;

        }

       


   



        /// <summary>
        /// Downloads a file from an FTP site
        /// </summary>
        /// <param name="sourceFileUrl">Remote file Url</param>
        /// <param name="destinationFilePath">Destination file path</param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <returns>Result of file download</returns>
        public static string DownloadFile(string sourceFileUrl, string destinationFilePath, string username = Constants.FTP.Username, string password = Constants.FTP.Password)
        {
            string output;

            // Get the object used to communicate with the server.
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(sourceFileUrl);

            //Specify the method of transaction
            request.Method = WebRequestMethods.Ftp.DownloadFile;

            // This example assumes the FTP site uses anonymous logon.
            request.Credentials = new NetworkCredential(username, password);

            //Indicate Binary so that any file type can be downloaded
            request.UseBinary = true;

            try
            {
                //Create an instance of a Response object
                using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
                {
                    //Request a Response from the server
                    using (Stream stream = response.GetResponseStream())
                    {
                        //Build a variable to hold the data using a size of 1Mb or 1024 bytes
                        byte[] buffer = new byte[1024]; //1 Mb chucks

                        //Establish a file stream to collect data from the response
                        using (FileStream fs = new FileStream(destinationFilePath, FileMode.Create))
                        {
                            //Read data from the stream at the rate of the size of the buffer
                            int ReadCount = stream.Read(buffer, 0, buffer.Length);

                            //Loop until the stream data is complete
                            while (ReadCount > 0)
                            {
                                //Write the data to the file
                                fs.Write(buffer, 0, ReadCount);

                                //Read data from the stream at the rate of the size of the buffer
                                ReadCount = stream.Read(buffer, 0, buffer.Length);
                            }
                        }
                    }

                    //Output the results to the return string
                    output = $"Download Complete, status {response.StatusDescription}";
                }

            }
            catch (Exception e)
            {
                //Something went wrong
                output = e.Message;
            }

            Thread.Sleep(Constants.FTP.OperationPauseTime);

            //Return the output of the Responce
            return (output);
        }

        public static void AggregateFunctions(List<Student> st)
           
        {

            Console.WriteLine("Count of all items in the List:::" + st.Count);
            List<int> age_list = new List<int>();
            List<Student> start_list = new List<Student>();
            List<Student> contains_list = new List<Student>();

            int startwith_count = 0;
            int contains_count = 0;
            foreach (var student in st)
            {
                Boolean s = student.MyRecord;
                int age = student.Age;
                age_list.Add(age);
                String enter_character= "C";

                if(student.FirstName.StartsWith(enter_character) || student.LastName.StartsWith(enter_character))
                {
                    startwith_count++;
                    start_list.Add(student);
                }


                if (student.FirstName.Contains(enter_character) || student.LastName.Contains(enter_character))
                {
                    contains_count++;
                    contains_list.Add(student);
                }

                if (s == true)
                {
                    Console.WriteLine("My Record::"+student);
                  
                  
                }
                
              }
            Console.WriteLine("Starts With C students count:" + start_list.Count);
            
            Console.WriteLine("Contains C students count:" + contains_list.Count);


            foreach(var student in start_list)
        {
                Console.Write("Students Start with C list::  "+ student.Directory + "  \n");
            }
            
            foreach (var student in contains_list)
            {
                Console.Write("Students Contains C list :: " + student.Directory + "  \n");
            }

            double average = age_list.Count > 0 ? age_list.Average() : 0;
            int min = age_list.Count > 0 ? age_list.Min() : 0;
            int max = age_list.Count > 0 ? age_list.Max() : 0;

           

            Console.WriteLine("Average Age:" + (int)Math.Round(average));

            Console.WriteLine("Min Age:" + min);
            Console.WriteLine("Max Age:" + max);


        }

        public static string ConvertCsvFileToJsonObject(string path)
        {
            var csv = new List<string[]>();
            var lines = File.ReadAllLines(path);

            foreach (string line in lines)
                csv.Add(line.Split(','));

            var properties = lines[0].Split(',');

            var listObjResult = new List<Dictionary<string, string>>();

            for (int i = 1; i < lines.Length; i++)
            {
                var objResult = new Dictionary<string, string>();
                for (int j = 0; j < properties.Length; j++)
                    objResult.Add(properties[j], csv[i][j]);

                listObjResult.Add(objResult);
            }

            return JsonConvert.SerializeObject(listObjResult);
        }


        }
       
}