using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace Signer.IsaHub.JsonClasses.Pdf
{
    
    internal class Result
    {

        public String result { get; set; }
        public int code { get; set; }
        public String message { get; set; }

        [JsonProperty(Required = Required.Default)]
        public byte[] FileData => Convert.FromBase64String(result??"");

    }
}
