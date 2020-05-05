using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Xps.Packaging;
using Microsoft.Office.Interop.Excel;

namespace ReadExcelData
{
	class ConvertToPDF
	{
		private Microsoft.Office.Interop.Excel.Application excelApp;
		private Microsoft.Office.Interop.Excel.Workbook excelWorkbook;


		public void Converter(string filePath)
		{
			try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.ScreenUpdating = false;
                excelApp.DisplayAlerts = false;
                ConvertExcelToPDFDoc(filePath);
				
			}
            catch { }
		}

		private void ConvertExcelToPDFDoc(string excelDocName)
		{
				excelDocName = excelDocName.Replace(new FileInfo(excelDocName).Extension, "") + ".xlsx";
				excelWorkbook = excelApp.Workbooks.Open(excelDocName);
				var pdfDocName = (new DirectoryInfo(excelDocName)).FullName;
                pdfDocName = pdfDocName.Replace(new FileInfo(excelDocName).Extension, "") + "";

                if (excelWorkbook == null)
                {
                    excelApp.Quit();

                    excelApp = null;
                    excelWorkbook = null;

                    return;
                }

                Worksheet sheet = excelWorkbook.Worksheets[1];
                sheet.PageSetup.TopMargin = 0;
                sheet.PageSetup.BottomMargin = 0;
                sheet.PageSetup.LeftMargin = 0;
                sheet.PageSetup.RightMargin = 0;
                sheet.PageSetup.FooterMargin = 0;
                sheet.PageSetup.HeaderMargin = 0;

				for (int i = 0; i < 2; i++)
			{
				try
				{
					sheet.PageSetup.FitToPagesTall = true;
					sheet.PageSetup.FitToPagesWide = true;
				}
                catch
				{ }
		}


                try
                {
                    excelWorkbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdfDocName);
                }
                catch
                {
                }
                finally
                {
                    excelWorkbook.Close();
                    excelApp.Quit();
                    excelApp = null;
                    excelWorkbook = null;
                }



		}
	}
}
