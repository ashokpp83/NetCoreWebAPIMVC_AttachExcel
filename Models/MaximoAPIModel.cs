using System;
using System.Collections.Generic;
using System.Text;

namespace WebAPIMVC_AttachExcel
{
    internal class MaximoAPIModel
    {
    }

    internal class SuccessRootObject
    {
        public List<PublicData_Member> member { get; set; }
        public string href { get; set; }
    }

    internal class PublicData_Member
    {

        public PublicData_Doclinks doclinks { get; set; }
        public int wonum { get; set; }
        public string siteid { get; set; }
    }

    internal class PublicData_Doclinks
    {
        public string href { get; set; }
    }


    internal class ErrorRootObject
    {
        public ErrorData Error { get; set; }
    }

    internal class ErrorData
    {
        public string message { get; set; }
        public string reasonCode { get; set; }
        public string statusCode { get; set; }
    }

}
