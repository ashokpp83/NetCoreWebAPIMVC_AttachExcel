using NLog;
using Spire.Xls;
using System;
using System.IO;
using System.Net;
using System.Text;

namespace WebAPIMVC_AttachExcel
{
    internal class MaximoAPIService
    {

        private static Logger logMe;

        private string mxAPIurl = string.Empty; 
        private string mxAPIKey = string.Empty; 
        
        private string workOrder = string.Empty;
        private string siteID = string.Empty;
        
        public string errMsg
        {
            get; set;
        }

        public MaximoAPIService(string mxWO, string mxSiteId, string _mxAPIurl, string _mxAPIKey, Logger _logMe)
        {
            workOrder = mxWO;
            siteID = mxSiteId;

            mxAPIurl = _mxAPIurl;
            mxAPIKey = _mxAPIKey;

            logMe = _logMe;
        }

        //Upload PDF to Maximo WO
        public void UploadExcelToMaximo(string mx_WO, string mx_SiteID, string filePath)
        {
            string url = string.Empty;
            try
            {
                logMe.Info("Uploading PDF Report : " + filePath + " to maximo WO: " + mx_WO);

                HttpStatusCode returnHTTPStatusCode = 0;

                //APi to get WO detail by wonum and Siteid
                url = mxAPIurl + @"api/os/mxwodetail?lean=1&apikey=" + mxAPIKey + "&oslc.select=*&oslc.where=wonum=" + mx_WO + " and siteid=" + (char)34 + mx_SiteID + (char)34; //(char)34 is for double quote

                //get href 
                string jsonResponse = GetMaximoAPIResponse(url, ref returnHTTPStatusCode);

                if (string.IsNullOrEmpty(jsonResponse))
                {
                    logMe.Error("Failed to get the Maximo API response for : " + url);
                    errMsg += Environment.NewLine + "Failed to get the Maximo API response for : " + url;
                }
                else
                {
                    //Deserialize and get doclinks API for WO and siteid
                    SuccessRootObject obj = Newtonsoft.Json.JsonConvert.DeserializeObject<SuccessRootObject>(jsonResponse);

                    //href + apikey  //"https://aamlxmaxapd001.aam.net:443/maximo/api/os/mxwodetail/_UkhUQy81Njk1OQ--/doclinks?apikey=" + apikey;
                    url = obj.member[0].doclinks.href + "?apikey=" + mxAPIKey;

                    //change it to api 
                    url = url.Replace("/oslc/", "/api/");

                    HttpStatusCode statusCode = UploadPDFReportToMaximoAPI(url, filePath);
                }
            }
            catch (Exception ex)
            {
                errMsg += Environment.NewLine + "Exception occured at UploadPDFToMaximo. file: " + filePath + ". Error: " + ex.Message;
                errMsg += Environment.NewLine + "Maximo url tried to upload: " + url;
                logMe.Error("Exception occured at UploadPDFToMaximo. file: " + "Exception occured at UploadPDFToMaximo. file: " + filePath + ". Error: " + ex.Message);
                logMe.Error("Maximo url tried to upload: " + url);
            }            
        }

       

        //attach the generated PDF Report to the given WorkOrder
        public HttpStatusCode UploadPDFReportToMaximoAPI(string url, string filePath)
        {
            HttpStatusCode statusCode = 0;

            //Create the HttpWebRequest object
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);

            //Set the timeout to 20 second
            req.Timeout = 20000;
            
            //enable accepting gzip content
            req.Method = "POST";
            //req.Accept = "*/*";

            try
            {
                req.Headers.Add("Content-Type: application/octet-stream");
                req.Headers.Add("Slug: " + Path.GetFileName(filePath));
                req.Headers.Add("X-document-meta: FILE/Attachments");
                req.Headers.Add("Accept-Encoding: Deflate,gzip");

                string desc = Path.GetFileName(filePath); //.Replace(Path.GetFileNameWithoutExtension(templateName) + "_", string.Empty);

                req.Headers.Add("X-document-description: " + desc);

                Stream stream = req.GetRequestStream();

                FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                byte[] buffer = new byte[4096];
                int bytesRead = 0;
                while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) != 0)
                    stream.Write(buffer, 0, bytesRead);
                fileStream.Close();

                stream.Close();

                using (HttpWebResponse webresponse = (HttpWebResponse)req.GetResponse())
                {
                    statusCode = webresponse.StatusCode;

                    webresponse.Close();
                }                
            }
            catch (WebException ex)
            {
                StreamReader reader = new StreamReader(ex.Response.GetResponseStream());

                errMsg += Environment.NewLine + "Exception occured at UploadPDFReportToMaximoAPI: url:" + url;
                errMsg += Environment.NewLine + "Error: " + reader.ReadToEnd();

                logMe.Error("Exception occured at UploadPDFReportToMaximoAPI while fetching the url: " + url + ". Error: " + reader.ReadToEnd());

                if (reader != null)
                    reader.Close();
            }
            catch (Exception ex)
            {
                errMsg += Environment.NewLine + "Exception occured at UploadPDFReportToMaximoAPI: " + ex.Message + ". while fetching the url: " + url;
                logMe.Error("Exception occured at UploadPDFReportToMaximoAPI: " + ex.Message + ". while fetching the url: " + url);
            }            

            return statusCode;
        }

       

        public string GetMaximoAPIResponse(string Url, ref HttpStatusCode returnHTTPStatusCode)
        {
            string returnData = string.Empty;

            HttpWebResponse webresponse = null;
            StreamReader reader = null;
            Stream responseStream = null;
            //Create the HttpWebRequest object
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(Url);

            //Set the timeout to 20 second
            req.Timeout = 20000;
            //req.Accept = "*/*";

            //enable accepting gzip content
            req.Method = "GET";
            req.Accept = "application/json";

            req.Headers.Add("Accept-Encoding: Deflate,gzip");
            req.Headers.Add("Content-Type: application/json");

            try
            {

                webresponse = (HttpWebResponse)req.GetResponse();

                // Get the stream associated with the response, just get the stream, but do not use it yet
                responseStream = webresponse.GetResponseStream();

                reader = new StreamReader(responseStream, Encoding.UTF8);
                returnData = reader.ReadToEnd();

                returnHTTPStatusCode = webresponse.StatusCode;

                webresponse.Close();
            }
            catch (WebException ex)
            {
                responseStream = ex.Response.GetResponseStream();

                reader = new StreamReader(responseStream);

                errMsg += Environment.NewLine + reader.ReadToEnd();
            }
            catch (Exception ex)
            {
                errMsg += Environment.NewLine + "Exception occured at GetMaximoAPIResponse: " + ex.Message + ". while fetching the url: " + Url;
                logMe.Error("Exception occured at GetMaximoAPIResponse: " + ex.Message + ". while fetching the url: " + Url);                
            }
            finally
            {
                if(reader != null)
                    reader.Close();
                if (responseStream != null)
                    responseStream.Close();
            }

            return returnData;
        }
    }
}
