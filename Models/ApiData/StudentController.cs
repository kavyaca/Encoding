using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using CSV.Models.Utilities;
using Newtonsoft.Json;

namespace CSV.Models.ApiData
{
    public class StudentController
    {

        public static void CallApi()
        {

            String url = Constants.Locations.api_url;
            RunTask(url);
        }



        public static async Task RunTask(String url)
        {

            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);

            httpWebRequest.Method = "GET";

            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();

            string jsonString = "";

            using (Stream stream = httpWebResponse.GetResponseStream())
            {
                StreamReader streamReader = new StreamReader(stream, System.Text.Encoding.UTF8);
                jsonString = streamReader.ReadToEnd();
            }
            List<StudentNewModel> st = JsonConvert.DeserializeObject<List<StudentNewModel>>(jsonString);


            foreach (var stud in st)
            {
                Console.WriteLine("Hi:: " + stud.Name);
            }

            WordDocumentClass.CreateWordprocessingDocumentFromApi(Constants.Locations.StudentDocFileFromApi, st);

        }
    }
}