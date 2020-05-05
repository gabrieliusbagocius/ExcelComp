using System.Data;
using System.IO;
using ExcelDataReader;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.CustomProperties;

namespace ReadExcelData
{
	public class BrainOfTheComparison
	{
		public void Program(string filePath1, string filePath2, string filePath3)
		{

			bool whichFileIsBeingProcessed = true;
			int countSequence = 0;
			string filePath = "";
			var workBook = new XLWorkbook();

			System.Data.DataTable newDataTable1 = null;
			System.Data.DataTable newDataTable2 = null;

			GetElements getElements = new GetElements();
			WorkWithStrings workWithString = new WorkWithStrings();
			CompareAndConvert compareAndConvert = new CompareAndConvert();
			FileStream stream;



			while (countSequence != 2)
			{
				if (whichFileIsBeingProcessed == true)
				{
					filePath = filePath1;
					whichFileIsBeingProcessed = false;
				}
				else
				{
					filePath = filePath2;
					whichFileIsBeingProcessed = true;
				}
				try
				{
					stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
						}
				catch
				{ return; }

					using (stream)
					{
						using (var reader = ExcelReaderFactory.CreateReader(stream))
						{
						do
						{
							while (reader.Read())
							{
							}
						} while (reader.NextResult());
						var result = reader.AsDataSet();
							System.Data.DataTable originalDataTable = result.Tables[0];
							var rowCollectionValue = originalDataTable.Rows;
							var amountOfRows = rowCollectionValue.Count - 1;
							int[] columnSetup = getElements.GetColumnElements(originalDataTable);
							originalDataTable = workWithString.ExtraLineRemoveLogic(result, amountOfRows, columnSetup);
							amountOfRows = originalDataTable.Rows.Count - 1;

							if (countSequence % 2 == 0)
							{
								newDataTable1 = originalDataTable;
							}
							else
							{
								newDataTable2 = originalDataTable;
							}
						}

					}

				if (countSequence == 1)
				{
					var finishedTable = compareAndConvert.DoEditThenCompare(newDataTable1, newDataTable2);
                    if (finishedTable != null)
                    {
                        var amountOfRows = finishedTable.Rows.Count;
                        finishedTable = FixColumnNames(finishedTable);
                        var setForExport = GetDataSetExportToExcel(finishedTable);
                        FinishAdditionalsAndExport toSetupAndExport = new FinishAdditionalsAndExport();
                        System.Data.DataTable tableForExport = setForExport.Tables[0];
                        var workSheet = workBook.Worksheets.Add(setForExport.Tables[0]);
                        var receivedSheet = toSetupAndExport.SetupAdditionals(workSheet, tableForExport);
                        toSetupAndExport.ExportToExcel(receivedSheet, workBook, filePath3);
                        ConvertToPDF convertToPDF = new ConvertToPDF();
                        convertToPDF.Converter(filePath3);
                    }
				}
	
				countSequence++;
			}
		}

		private DataTable FixColumnNames(DataTable finishedTable)
		{
			finishedTable.Columns[0].ColumnName = finishedTable.Rows[0][0].ToString();
			finishedTable.Columns[1].ColumnName = finishedTable.Rows[0][1].ToString();
			finishedTable.Columns[2].ColumnName = finishedTable.Rows[0][2].ToString();
			finishedTable.Columns[3].ColumnName = finishedTable.Rows[0][3].ToString();
			finishedTable.Columns[4].ColumnName = finishedTable.Rows[0][4].ToString();
			finishedTable.Columns[5].ColumnName = finishedTable.Rows[0][5].ToString();
			finishedTable.Columns[6].ColumnName = finishedTable.Rows[0][6].ToString();
			finishedTable.Rows[0].Delete();
			return finishedTable;
		}


		private static DataSet GetDataSetExportToExcel(System.Data.DataTable resultTable)
		{
			DataSet resultDataSet = new DataSet();
			System.Data.DataTable dataTemp = new System.Data.DataTable();
			dataTemp = resultTable;
			resultDataSet.Tables.Add(dataTemp);
			resultDataSet.Tables[0].TableName = "Compared";
			return resultDataSet;
		}
	}

}
