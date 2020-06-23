using Newtonsoft.Json;
using NLog;
using WebAPIMVC_AttachExcel.Interfaces;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net;
using System.Text;
using Spire.Xls;

namespace WebAPIMVC_AttachExcel
{
    //AttachExcelService class
    public class AttachExcelService : IAttachExcelService
    {
        private static readonly Logger logMe = LogManager.GetCurrentClassLogger();

        private string excel_FilePath = string.Empty;

        private string maximoWorkOrder = string.Empty;
        private string maximoSiteID = string.Empty;
        private string excelFileName_FromTemplate = string.Empty;

        private string maximoAPI_url = string.Empty; 
        private string maximo_APIkey = string.Empty;

        private string errMsg = string.Empty;
        private string emailSubject = "Error in AttachExcelService. " + " (" + Environment.MachineName + " - " + ")";
        private IEmailService emailservice;

        //AttachExcelService constructor
        public AttachExcelService(Dictionary<string, string> AppSettings, IEmailService emailSvc)
        {
            excel_FilePath = AppSettings["Excel_FilePath"];

            maximoWorkOrder = AppSettings["ParsableName_MaximoWorkOrder"];
            maximoSiteID = AppSettings["ParsableName_MaximoSiteID"];
            excelFileName_FromTemplate = AppSettings["ExcelFileName_FromTemplate"];
            

            maximoAPI_url = AppSettings["MaximoAPI_url"];
            maximo_APIkey = AppSettings["Maximo_APIkey"];

            emailservice = emailSvc;

        }

        //AttachExcelToMaximo
        public bool AttachExcelToMaximo(string json)
        {  
            bool success = false;

            try
            {
                logMe.Info("Received parsableJobData {0}", json);

                Dictionary<string, object> kvp = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                
                //Maximo WorkOrder
                if (!kvp.ContainsKey(maximoWorkOrder))
                {
                    errMsg += "The key '" + maximoWorkOrder + "' is not provided in the parsable input data";
                    return success;
                }
                string mx_WO = kvp[maximoWorkOrder].ToString();

                //MaximoSiteId
                if (!kvp.ContainsKey(maximoSiteID))
                {
                    errMsg += "The key '" + maximoSiteID + "' is not provided in the parsable input data";
                    return success;
                }
                string mx_SiteID = kvp[maximoSiteID].ToString();

                if (!kvp.ContainsKey(excelFileName_FromTemplate))
                {
                    errMsg += "The key '" + excelFileName_FromTemplate + "' is not provided in the parsable input data";
                    return success;
                }
                string excelFileName = kvp[excelFileName_FromTemplate].ToString();

                //fullpath
                excel_FilePath = excel_FilePath + "\\" + excelFileName;

                //PDF Report fileFullPath to download
                string newfilePath = Path.GetDirectoryName(excel_FilePath) + "\\" + Path.GetFileNameWithoutExtension(excelFileName) + "_" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".xlsx";


                //save as
                if (Path.GetExtension(excelFileName) == ".xls" || Path.GetExtension(excelFileName) == ".xlsm")
                    ConvertXLSMToXLSX(excel_FilePath, newfilePath);
                else
                {
                    using (var app = new OfficeOpenXml.ExcelPackage(new FileInfo(excel_FilePath)))
                    {
                        //process
                        app.SaveAs(new FileInfo(newfilePath));
                    }
                }

                emailSubject = "Error while Uploading the PDF Report to Maximo Work Order" + " (" + Environment.MachineName + " - " + ")";
                MaximoAPIService mx_APIService = new MaximoAPIService(mx_WO, mx_SiteID, maximoAPI_url, maximo_APIkey, logMe);

                //Upload the PDF Report to Maximo WorkOrder                
                mx_APIService.UploadExcelToMaximo(mx_WO, mx_SiteID, newfilePath);

                if (!string.IsNullOrEmpty(mx_APIService.errMsg))
                {
                    errMsg += mx_APIService.errMsg;
                    //has error
                    return success;
                }

                success = true;
            }
            catch (Exception ex)
            {
                errMsg += "Exception occured in AttachExcelToMaximo. Error: " + ex.Message;
                logMe.Error("Exception occured in AttachExcelToMaximo. Error: " + ex.Message);
            }
            finally
            {
                if (!string.IsNullOrEmpty(errMsg))
                {
                    //send an Email to support team
                    emailservice.SendEmailNotification(emailSubject, errMsg);
                }
            }

            return success;
        }

        private void ConvertXLSMToXLSX(String sourceFile, string destFile)
        {
            Workbook workbook = new Workbook();
            try
            {
                workbook.LoadFromFile(sourceFile);
                workbook.SaveToFile(destFile, ExcelVersion.Version2016);
            }
            catch (Exception ex)
            {
                errMsg += Environment.NewLine + "Exception occured at ConvertXLSToXLSX: " + ex.Message;
                logMe.Error("Exception occured at ConvertXLSToXLSX: " + ex.Message);
            }
            finally
            {
                workbook.Dispose();
            }
        }
    }
}
