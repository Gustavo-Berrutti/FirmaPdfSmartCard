using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FirmaPdfWpf.ViewModels.FirmaDesdeCarpeta.Models
{
    public class EstadoFirmaArchivo
    {
        public String Nombre { get; private set; }
        public String Descripcion { get; set; }
        public String BackgroundColor { get; private set; }
        public String ForegroundColor { get; private set; }

        public Boolean ColorDefined {
            get {
                return !String.IsNullOrWhiteSpace(BackgroundColor);
            }
        }

        private EstadoFirmaArchivo(string nombre, string descripcion, string backgroundColor, string foregroundColor)
        {
            Nombre = nombre;
            Descripcion = descripcion;
            BackgroundColor = backgroundColor;
            ForegroundColor = foregroundColor;
        }

        public static EstadoFirmaArchivo GetPendiente() {
            return new EstadoFirmaArchivo("Pendiente", "Esperando Ejecución", "", "");
        }

        public static EstadoFirmaArchivo GetError()
        {
            return new EstadoFirmaArchivo("Error", "Error firmando el archivo", "#c0392b", "Whitesmoke");
        }
        public static EstadoFirmaArchivo GetAtencion()

        {
            return new EstadoFirmaArchivo("Atención", "Archvio firmado, pero ocurrió otro problema", "#f1c40f", "Gray");
        }

        public static EstadoFirmaArchivo GetFirmado()
        {
            return new EstadoFirmaArchivo("Firmado", "El archivo ha sido firmado", "#27ae60", "Whitesmoke");
        }

    }
}
