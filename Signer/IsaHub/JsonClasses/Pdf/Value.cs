using System;
using System.Collections.Generic;
using System.Text;

namespace Signer.IsaHub.JsonClasses.Pdf
{
    internal class Value
    {
        public Value()
        {
            config = new Config();

        }
        public Config config { get; set; }
        public String content { get; set; }
    }
}
