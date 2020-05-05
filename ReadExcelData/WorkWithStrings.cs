using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

namespace ReadExcelData
{
	public class WorkWithStrings
	{
		public DataTable ExtraLineRemoveLogic(DataSet tableResult, int amountOfRows, int[] tempColumnSetup)
		{
			System.Data.DataTable originalDataTable = tableResult.Tables[0];
			String stringModifier = String.Empty;
			var tempOriginalArray = originalDataTable;
			var initializeEditThenCompare = new CompareAndConvert();
			List<string> allTableConverted = new List<string>();
			List<Int32> valueToCheck = new List<Int32>();



			for (int i = 0; i <= amountOfRows; i++)
			{
				allTableConverted.Add(originalDataTable.Rows[i][tempColumnSetup[0]].ToString());
				allTableConverted.Add(originalDataTable.Rows[i][tempColumnSetup[1]].ToString());
				allTableConverted.Add(originalDataTable.Rows[i][tempColumnSetup[2]].ToString());
				allTableConverted.Add(originalDataTable.Rows[i][tempColumnSetup[3]].ToString());
				allTableConverted.Add(originalDataTable.Rows[i][tempColumnSetup[4]].ToString());
				allTableConverted.Add(originalDataTable.Rows[i][tempColumnSetup[5]].ToString());
				allTableConverted.Add(originalDataTable.Rows[i][tempColumnSetup[6]].ToString());


				if ( string.IsNullOrEmpty(allTableConverted[6]) != true)
				{
					valueToCheck.Add(allTableConverted[0].Split('\n').Length);
					valueToCheck.Add(allTableConverted[1].Split('\n').Length);
					valueToCheck.Add(allTableConverted[2].Split('\n').Length);
					valueToCheck.Add(allTableConverted[3].Split('\n').Length);
					valueToCheck.Add(allTableConverted[4].Split('\n').Length);
					valueToCheck.Add(allTableConverted[5].Split('\n').Length);
					valueToCheck.Add(allTableConverted[6].Split('\n').Length);
					int countValueToCheck = valueToCheck.Count();

					for (int p = 0; p < countValueToCheck; p++)
					{
						if (valueToCheck[p] == 1)
						{
							tempOriginalArray.Rows[i][tempColumnSetup[p]] = allTableConverted[p];
						}
						if (valueToCheck[p] == 2)
						{
							string newOutputString = StringLineRemoval(allTableConverted[p], valueToCheck[p], p);
							tempOriginalArray.Rows[i][tempColumnSetup[p]] = newOutputString;
						}
						if (valueToCheck[p] == 3)
						{
							string newOutputString = StringLineRemoval(allTableConverted[p], valueToCheck[p], p);
							tempOriginalArray.Rows[i][tempColumnSetup[p]] = newOutputString;
						}

					}

					allTableConverted.Clear();
				}
				valueToCheck.Clear();
			}
			DataTable finishedDataTable = RemoveDuplicateRows(tempOriginalArray, tableResult, tempColumnSetup);
			return finishedDataTable;
		}

		private string StringLineRemoval(string tableResult, int count, int clmnNumber)
		{
			string outputString = "";
			int tempNumber = 1;
			string tempString = "";

			if (count == 1)
			{
				string[] linesValue = tableResult.Split(Environment.NewLine.ToCharArray()).Skip(1).ToArray();
				outputString = string.Join("", linesValue);
			}

			if (count == 2)
			{
				string[] linesValue = tableResult.Split(Environment.NewLine.ToCharArray()).Skip(1).ToArray();
				outputString = string.Join("", linesValue);

			}
			
			if (count == 3)
			{

				string[] linesValue = tableResult.Split(Environment.NewLine.ToCharArray()).Skip(1).ToArray();
				string linesValueConverted = string.Join(Environment.NewLine, linesValue);

				string newValue = "";
				if (linesValueConverted.Contains("\r\n"))
					newValue = linesValueConverted.Substring(0, linesValueConverted.IndexOf("\r\n"));
				else
					newValue = linesValueConverted;
				string[] mainLinesValue = linesValueConverted.Split(Environment.NewLine.ToCharArray()).Skip(1).ToArray();
				string mainLinesValueConverted = string.Join("", mainLinesValue);
				
				if (clmnNumber == 2)
				{
					if (string.IsNullOrEmpty(newValue))
					{
						outputString = mainLinesValueConverted;
					}
					else
					{
						if (string.IsNullOrEmpty(mainLinesValueConverted))
						{
							outputString = newValue;
						}
						else
						{
							outputString = mainLinesValueConverted + ", " + newValue;
						}
					}
				}
				else
				{
					outputString = newValue + " " + mainLinesValueConverted;

				}
			}
			return outputString;
		}

		private DataTable RemoveDuplicateRows(DataTable receivedDataTable, DataSet tableResult, int[] columnSetup)
		{
			int columnNumber = columnSetup[1];
			Hashtable hashTable = new Hashtable();
			ArrayList duplicateList = new ArrayList();
			DataRow[] rows = receivedDataTable.Rows.Cast<DataRow>().ToArray();
			Array.Reverse(rows);
			foreach (DataRow drow in rows)
			{
				if (hashTable.Contains(drow[columnNumber]))
				{
					duplicateList.Add(drow);
				}
				else
				{
					hashTable.Add(drow[columnNumber], string.Empty);
				}
			}
			foreach (DataRow dRow in duplicateList)
			{
				receivedDataTable.Rows.Remove(dRow);
			}
			return receivedDataTable;
		}

		public DataTable RemoveEmptyRows(DataTable receivedDataTable)
		{
			for (int i = receivedDataTable.Rows.Count; i >= 1; i--)
			{
				DataRow currentRow = receivedDataTable.Rows[i - 1];
				foreach (var colValue in currentRow.ItemArray)
				{
					if (!string.IsNullOrEmpty(colValue.ToString()))
						break;
					receivedDataTable.Rows[i - 1].Delete();
				}
			}
			return receivedDataTable;
		}
	}
}