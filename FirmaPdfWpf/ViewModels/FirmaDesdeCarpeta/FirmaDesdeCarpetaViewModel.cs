using FirmaPdfWpf.ViewModels.FirmaDesdeCarpeta.Models;
using Serilog;
using Serilog.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using Signer;
using Signer.Impl;
using MahApps.Metro.Controls.Dialogs;
using Signer.Pdf;

namespace FirmaPdfWpf.ViewModels.FirmaDesdeCarpeta
{
    public class FirmaDesdeCarpetaViewModel : ViewModelBase
    {
        ILogger logger = Log.ForContext<FirmaDesdeCarpetaViewModel>();


        private String carpetaOrigen;
        private String carpetaDestino;

        public FirmaDesdeCarpetaViewModel()
        {
            ArchivosParaFirmar = new ObservableCollection<FileProcessing>();

        }

        public String CarpetaOrigen
        {
            get { return carpetaOrigen; }
            set
            {
                carpetaOrigen = value;
                OnPropertyChanged();
            }
        }
        public String CarpetaDestino
        {
            get { return carpetaDestino; }
            set
            {
                carpetaDestino = value;
                OnPropertyChanged();
            }
        }


        public ObservableCollection<FileProcessing> ArchivosParaFirmar
        {
            get; private set;
        }


        #region  Métodos

        public bool LoadFiles()
        {
            var dirOrigen = new DirectoryInfo(CarpetaOrigen);
            var files = dirOrigen.GetFiles("*.pdf", SearchOption.AllDirectories);
            if (files.Count() == 0)
                return false;

            foreach (var f in files)
            {
                ArchivosParaFirmar.Add(new FileProcessing()
                {
                    File = f,
                    Estado = EstadoFirmaArchivo.GetPendiente(),
                    Lote = f.Directory.Name,
                });
            }
            return true;
        }

        public string ValidateData()
        {
            StringBuilder sb = new StringBuilder();
            var dirOrigen = new DirectoryInfo(CarpetaOrigen);
            if (!dirOrigen.Exists)
                sb.AppendLine("** El directorio origen no existe.");
            var dirDestino = new DirectoryInfo(CarpetaDestino);
            if (!dirDestino.Exists)
                sb.AppendLine("** El directorio destino no existe.");
            return sb.ToString();

        }

        public void ProcessFiles(PdfSignerBase pdfSigner)
        {
            pdfSigner.SignedFileEvent += PdfSigner_SignedFileEvent;

            foreach (var a in ArchivosParaFirmar)
            {
                logger.Information("Firmando archivo {archivo} del lote {lote}", a.File.Name, a.Lote);
                pdfSigner.SignPdf(a.File);
            }

        }
        #endregion


        #region  Event handlers

        private void PdfSigner_SignedFileEvent(object sender, SignedFileEventArgs e)
        {
            var a = ArchivosParaFirmar.Where(ar => ar.File.FullName == e.File.FullName).First();

            if (e.HasError)
            {
                a.Estado = EstadoFirmaArchivo.GetError();
                a.Estado.Descripcion = e.Message;
                logger.Error("Ha ocurrido un error firmando el archivo {archivo} del lote {lote} con el siguiente mensaje: {mensaje}", e.File.Name, a.Lote, e.Message);
            }
            else
            {
                try
                {
                    File.WriteAllBytes(System.IO.Path.Combine(carpetaDestino, e.File.Name), e.SignedContent);
                }
                catch (Exception es)
                {
                    logger.Error("No se pudo guardar el archivo {directorio} .", a.File.FullName);

                    a.Estado = EstadoFirmaArchivo.GetError();
                    a.Estado.Descripcion = $"No se pudo guardar el archivo.\n{es.Message}";
                    return;
                }
                a.Estado = EstadoFirmaArchivo.GetFirmado();
                logger.Information("El archivo {archivo} del lote {lote} ha sido firmado.", a.File.Name, a.Lote);

                try
                {
                    a.File.Delete();
                    logger.Information("El archivo {archivo} del lote {lote} ha sido eliminado de la carpeta origen.", a.File.Name, a.Lote);
                    if (a.File.Directory.GetFiles().Count() == 0)
                        try
                        {
                            a.File.Directory.Delete();
                            logger.Information("Se borró el directorio {directorio} de la carpeta origen.", a.File.Directory.Name);
                        }
                        catch (Exception eD)
                        {
                            logger.Error(eD, "No se pudo borrar el directorio {directorio} de la carpeta origen.", a.File.Directory.Name);
                            a.Estado = EstadoFirmaArchivo.GetAtencion();
                            a.Estado.Descripcion = $"El archivo se firmó y se eliminó.\nEra el último archivo del lote pero no se pudo borrar el directorio.\n{eD.Message}";

                        }
                }
                catch (Exception exx)
                {
                    logger.Error(exx, "No se pudo borrar archivo {archivo} del lote {lote} de la carpeta origen.", a.File.Name, a.Lote);
                    a.Estado = EstadoFirmaArchivo.GetAtencion();
                    a.Estado.Descripcion = $"El archivo se firmó pero no se pudo eliminar.\n{exx.Message}";
                }
            }

        }
        #endregion
    }
}
