using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FirmaPdfWpf.ViewModels.FirmaDesdeCarpeta.Models
{
    public class FileProcessing : ViewModelBase
    {
        private string lote;
        private FileInfo file;
        private EstadoFirmaArchivo estado;

        public String Lote
        {
            get => lote;
            set
            {
                lote = value; 
                OnPropertyChanged();
            }
        }

        public FileInfo File
        {
            get => file;
            set
            {
                file = value;
                OnPropertyChanged();
            }
        }

        public EstadoFirmaArchivo Estado
        {
            get => estado; set
            {
                estado = value;
               OnPropertyChanged();
            }
        }

    }
}
