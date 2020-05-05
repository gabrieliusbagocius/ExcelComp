using System.Linq;
using ClosedXML.Excel;
using System.Data;
using System.IO;

namespace ReadExcelData
{
	public class FinishAdditionalsAndExport
	{
		public IXLWorksheet SetupAdditionals(IXLWorksheet workSheet, System.Data.DataTable tableForExport)
		{
			var tableRowsResult = tableForExport.Rows.Cast<DataRow>().ToArray().Count();
			var tableColumnsResult = tableForExport.Columns.Cast<DataColumn>().ToArray().Count();
			SetupFirstStaticColumns(workSheet);
			string someString = "";
			int positionOfNewLine = 0;
            int lengthOfNewString = 0;
            int tempStringLength = 0;
            string copyOfSomeString;
            int secondPositionOfNewLine = 0;
            int finalStringLength = 0;
            int firstIndex = -1;
            int secondIndex = -1;


            for (int i = 0; i < tableColumnsResult; i++)
			{
				for (int p = 0; p < tableRowsResult; p++)
				{
					someString = tableForExport.Rows[p][i].ToString();
                    copyOfSomeString = someString;
					positionOfNewLine = someString.IndexOf("\r\n");
                    tempStringLength = someString.Count();

                    if (positionOfNewLine >= 0)
                    {
                        firstIndex = positionOfNewLine;
                         if (positionOfNewLine != 0)
						{
							workSheet.Cell(p + 2, i + 1).RichText.Substring(0, firstIndex).SetFontColor(XLColor.Red).SetStrikethrough();
						}
                            copyOfSomeString = copyOfSomeString.Remove(0, positionOfNewLine + 2);
                            secondPositionOfNewLine = copyOfSomeString.IndexOf("\r\n");
                            if (secondPositionOfNewLine > 0)
                            {
                                copyOfSomeString = copyOfSomeString.Remove(0, secondPositionOfNewLine + 2);
                                secondIndex = positionOfNewLine + 2 + secondPositionOfNewLine + 2;
                                lengthOfNewString = secondIndex - firstIndex;
                                finalStringLength = tempStringLength - secondIndex;
                                workSheet.Cell(p + 2, i + 1).RichText.Substring(firstIndex, lengthOfNewString).SetFontColor(XLColor.Green);
                            if (finalStringLength != 0)
                            {
                                workSheet.Cell(p + 2, i + 1).RichText.Substring(secondIndex, finalStringLength).SetFontColor(XLColor.Black);
                            }
                        }
                    }
                    SetupTheBorders(p + 2, i + 1, workSheet);

                }
                lengthOfNewString = 0;
                finalStringLength = 0;
            }
            return workSheet;
		}


        int IndexOfSecond(string theString, string toFind)
        {
            int first = theString.IndexOf(toFind);
            if (first == -1) return -1;
            return theString.IndexOf(toFind, first + 1);
        }

        private static void SetupTheBorders(int index1, int index2, IXLWorksheet workSheet)
		{
			workSheet.Cell(index1, index2).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
			workSheet.Cell(index1, index2).Style.Border.BottomBorderColor = XLColor.Black;

			workSheet.Cell(index1, index2).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
			workSheet.Cell(index1, index2).Style.Border.LeftBorderColor = XLColor.Black;

			workSheet.Cell(index1, index2).Style.Border.TopBorder = XLBorderStyleValues.Thin;
			workSheet.Cell(index1, index2).Style.Border.TopBorderColor = XLColor.Black;

			workSheet.Cell(index1, index2).Style.Border.RightBorder = XLBorderStyleValues.Thin;
			workSheet.Cell(index1, index2).Style.Border.RightBorderColor = XLColor.Black;

			workSheet.Cell(index1, index2).Style.Border.RightBorder = XLBorderStyleValues.Thin;
			workSheet.Cell(index1, index2).Style.Border.RightBorderColor = XLColor.Black;

			workSheet.Cell(index1, index2).Style.Border.RightBorder = XLBorderStyleValues.Thin;
			workSheet.Cell(index1, index2).Style.Border.RightBorderColor = XLColor.Black;
		}

		private static void SetupFirstStaticColumns(IXLWorksheet workSheet)
		{
			workSheet.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.Gray;
			SetupTheBorders(1, 1, workSheet);
			workSheet.Cell(1, 2).Style.Fill.BackgroundColor = XLColor.Gray;
			SetupTheBorders(1, 2, workSheet);
			workSheet.Cell(1, 3).Style.Fill.BackgroundColor = XLColor.Gray;
			SetupTheBorders(1, 3, workSheet);
			workSheet.Cell(1, 4).Style.Fill.BackgroundColor = XLColor.Gray;
			SetupTheBorders(1, 4, workSheet);
			workSheet.Cell(1, 5).Style.Fill.BackgroundColor = XLColor.Gray;
			SetupTheBorders(1, 5, workSheet);
			workSheet.Cell(1, 6).Style.Fill.BackgroundColor = XLColor.Gray;
			SetupTheBorders(1, 6, workSheet);
			workSheet.Cell(1, 7).Style.Fill.BackgroundColor = XLColor.Gray;
			SetupTheBorders(1, 7, workSheet);
		}
	
		public void ExportToExcel(IXLWorksheet sheet, XLWorkbook workBook, string filePath3)
		{
            try
            {
				filePath3 = filePath3.Replace(new FileInfo(filePath3).Extension, "") + ".xlsx";
				if (File.Exists(filePath3)) File.Delete(filePath3);
                workBook.SaveAs(filePath3);
                workBook.Dispose();
            }
            catch { }
		}
	}
}
