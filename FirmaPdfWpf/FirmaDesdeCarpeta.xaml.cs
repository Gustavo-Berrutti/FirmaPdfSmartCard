using Signer;
using Signer.Impl;
using Signer.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using MahApps.Metro.Controls;
using FirmaPdfWpf.ViewModels.FirmaDesdeCarpeta;
using MahApps.Metro.Controls.Dialogs;
using FirmaPdfWpf.ViewModels.FirmaDesdeCarpeta.Models;
using System.Collections.ObjectModel;
using Serilog;
using FirmaPdfWpf.Npoi;
using Microsoft.Win32;
using FirmaPdfWpf.ViewModels.Pin;

namespace FirmaPdfWpf
{
    /// <summary>
    /// Lógica de interacción para FirmaDesdeCarpeta.xaml
    /// </summary>
    public partial class FirmaDesdeCarpeta : MetroWindow
    {

        ILogger logger = Log.ForContext<FirmaDesdeCarpeta>();

        FirmaDesdeCarpetaViewModel viewModel = new FirmaDesdeCarpetaViewModel();


        public FirmaDesdeCarpeta()
        {
            InitializeComponent();
            DataContext = viewModel;

        }


        private void FrmFirmadesdeCarpeta_Load(object sender, RoutedEventArgs e)
        {
            viewModel.CarpetaOrigen = System.IO.Path.Combine(FirmaPdfWpf.Properties.Settings.Default.CarpetaOrigen, Environment.UserName);
            viewModel.CarpetaDestino = FirmaPdfWpf.Properties.Settings.Default.CarpetaDestino;


            logger.Information("Aplicación inicializada.");
        }

        private async void BtnFirmar_Click(object sender, RoutedEventArgs e)
        {
            logger.Information("Comienzo de Firma de documentos.");
            BtnFirmar.IsEnabled = false;
            BtnExportar.IsEnabled = false;


            try
            {
                String errores = viewModel.ValidateData();
                if (!String.IsNullOrWhiteSpace(errores))
                {
                    await this.ShowMessageAsync("Error", errores);
                    BtnFirmar.IsEnabled = true;
                    return;
                }
                viewModel.ArchivosParaFirmar.Clear();
                if (viewModel.LoadFiles())
                {
                    using (var pdfSigner = FirmaDocumentosFactory.GetInstance().getPdfSigner())
                    {
                        var run = true;
                        if (pdfSigner.RequiresPin)
                        {
                            var result = AskForPin();
                            if (result.dialogResultOk)
                            {
                                if (String.IsNullOrWhiteSpace(result.pin))
                                {
                                    await this.ShowMessageAsync("Error", "El pin no puede ser vacío.");
                                    run = false;
                                }
                                else
                                {
                                    pdfSigner.Pin = result.pin;
                                }

                            }
                            else
                                run = false;
                        }
                        if (run)
                            await Task.Run(() => viewModel.ProcessFiles(pdfSigner));
                    }
                }
                else
                {
                    await this.ShowMessageAsync("Información", "No hay Archivos para firmar.");
                    logger.Information("No se encontraron archivos para firmar.");
                }
            }

            catch (Exception ex)
            {
                logger.Error(ex, "Ha ocurrido un error inesperado firmando los archivos.");
                await this.ShowMessageAsync("Error", "Ocurrió un error inesperado.\n" + ex.Message);

            }

            BtnFirmar.IsEnabled = true;
            BtnExportar.IsEnabled = viewModel.ArchivosParaFirmar.Count > 0;
            this.Cursor = Cursors.Arrow;
        }



        private (Boolean dialogResultOk, string pin) AskForPin()
        {

            var pinModel = new PinViewModel();
            this.Cursor = Cursors.Wait;
            var w = new Pin(pinModel);
            w.Owner = this;
            var result = w.ShowDialog();
            return (result.HasValue ? result.Value : false, pinModel.Pin);
        }


        private async void BtnSalir_Click(object sender, RoutedEventArgs e)
        {
            MetroDialogSettings settings = new MetroDialogSettings() { AffirmativeButtonText = "Salir", NegativeButtonText = "Cancelar" };
            var result = await this.ShowMessageAsync("Atención", "Está a punto de salir de la aplicación\nSi el proceso de firma está en ejecución, se cortará inmediatamente.\n¿Desea continuar?", MessageDialogStyle.AffirmativeAndNegative, settings);
            if (result == MessageDialogResult.Affirmative)
                Application.Current.Shutdown();
        }

        private async void BtnExportar_Click(object sender, RoutedEventArgs e)
        {
            if (viewModel.ArchivosParaFirmar.Count > 0)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog()
                {
                    DefaultExt = "xls",
                    CheckPathExists = true,
                    Filter = "Archivos excel|*.xls",
                    Title = "Guardar Resumen de firma.",
                    ValidateNames = true,
                    AddExtension = true,
                    OverwritePrompt = true,
                    FileName = $"FirmaPdfs_{DateTime.Now.ToString("yyyyMMddHHmmss")}"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    var excExport = new ArchivosProcesadosExcelGenerator();
                    File.WriteAllBytes(saveFileDialog.FileName, excExport.GetReport(viewModel.ArchivosParaFirmar.ToList()));
                }
            }
            else
                await this.ShowMessageAsync("Información", "La grilla no contiene registros.");

        }
    }
}
