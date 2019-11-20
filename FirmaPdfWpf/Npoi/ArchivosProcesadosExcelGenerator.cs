using FirmaPdfWpf.ViewModels.FirmaDesdeCarpeta.Models;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FirmaPdfWpf.Npoi
{
    public class ArchivosProcesadosExcelGenerator : ExcelGeneratorBase
    {
        public byte[] GetReport(List<FileProcessing> data)
        {
            CreateWorkbook("Firma de Archivos");
            InitFontsStyles();
            var sheet = wb.CreateSheet("Archivos Firmados");

            var row = sheet.CreateRow(0);
            AddTitleCell(row, 0, 5, "Resultado de la firma de archivos");

            var rownumber = 2;
            addTablaArchivosFirmados(ref rownumber, data, sheet);

            return WriteToStream().GetBuffer();
        }

        private void addTablaArchivosFirmados(ref int rowNumber, List<FileProcessing> data, ISheet sheet)
        {
            var row = sheet.CreateRow(rowNumber);

            sheet.CreateFreezePane(0, rowNumber + 1);

            var cell = row.CreateCell(0);

            AddHeaderCell(row, 0, "Lote", widthInChars(25));
            AddHeaderCell(row, 1, "Archivo", widthInChars(55));
            AddHeaderCell(row, 2, "Estado", widthInChars(12));
            AddHeaderCell(row, 3, "Descripcion", widthInChars(32));
            AddHeaderCell(row, 4, "Path completo", widthInChars(92));
            rowNumber++;
            var firstRow = rowNumber;
            bool alternate = false;
            foreach (var pp in data)
            {
                row = sheet.CreateRow(rowNumber);
                addRow(row, alternate, pp);
                rowNumber++;
                alternate = !alternate;
            }

        }


        private void addRow(IRow row, bool alternate, FileProcessing file)
        {

            var cell = row.CreateCell(0);
            cell.SetCellValue(file.Lote);
            cell.CellStyle = alternate ? alternateTextStyle : textStyle;

            cell = row.CreateCell(1);
            cell.SetCellValue(file.File.Name);
            cell.CellStyle = alternate ? alternateTextStyle : textStyle;

            cell = row.CreateCell(2);
            cell.SetCellValue(file.Estado.Nombre);
            cell.CellStyle = alternate ? alternateTextStyle : textStyle;


            cell = row.CreateCell(3);
             cell.SetCellValue(file.Estado.Descripcion);
            cell.CellStyle = alternate ? alternateTextStyle : textStyle;

            cell = row.CreateCell(4);
            cell.SetCellValue(file.File.FullName);
            cell.CellStyle = alternate ? alternateTextStyle : textStyle;

        }

    }
}
