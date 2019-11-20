using System;
using System.Collections.Generic;
using System.Text;

namespace Signer.IsaHub.JsonClasses.Pdf
{
    internal class Config
    {
        public bool apariencia { get; set; }
        public String urlImagen { get; set; }
        public int paginaFirma { get; set; }
        public int modoFirma { get; set; }
        public List<int> posicion { get; set; }
    }
}
