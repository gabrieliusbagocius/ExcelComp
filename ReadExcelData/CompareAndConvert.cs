using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;

namespace ReadExcelData
{
	public class CompareAndConvert
	{
		public DataTable DoEditThenCompare(DataTable table1Data, DataTable table2Data)
		{
            bool isNull = false;
			GetElements getElements = new GetElements();
			DataRow[] table1Rows = table1Data.Rows.Cast<DataRow>().ToArray();
			DataRow[] table2Rows = table2Data.Rows.Cast<DataRow>().ToArray();
			int[] columnNumber1 = getElements.GetColumnElements(table1Data);
			int[] columnNumber2 = getElements.GetColumnElements(table2Data);


            int table1PartNumber = columnNumber1[6];
			int table2PartNumber = columnNumber2[6];
			List<string> table1ColumnList = ColumnConvertedToList(table1Rows, table1PartNumber);
			List<string> table2ColumnList = ColumnConvertedToList(table2Rows, table2PartNumber);


			IEnumerable<string> similiarityQuery =
			table1ColumnList.Intersect(table2ColumnList);
			var similarityCount = similiarityQuery.Count();


			int columnsIndex = 0;
			int[] newColumnSetup = new int[] { 0, 1, 2, 3, 4, 5, 6};
			var comparedData = new System.Data.DataTable();
			int columns = newColumnSetup.Count();
			int comparedRows = similarityCount;
			string[] comparedColumns = new string[columns];
			for (int i = 0; i < columns; i++)
				comparedData.Columns.Add();
			for (int i = 0; i < comparedRows; i++)
				comparedData.Rows.Add(comparedColumns);


			List<string> newList1 = new List<string> { };
			List<string> newList2 = new List<string> { };
			string compList = "";
			int tempForComparedData = 0;
			int lengthOfTable1 = table1ColumnList.Count();
			int lengthOfTable2 = table2ColumnList.Count();

            int countFirstDuplicateZeros = 0;
            int countSecondDuplicateZeros = 0;
             

            for (int i = 0; i < 7; i++)
            {
                if (columnNumber1[i] == 0)
                {
                    countFirstDuplicateZeros++;
                }
                if (columnNumber2[i] == 0)
                {
                    countSecondDuplicateZeros++;
                }
                if (countFirstDuplicateZeros > 1 || countSecondDuplicateZeros > 1)
                {
                    isNull = true;
                    comparedData = null;
                    return comparedData;
                }
            }


            for (int k = 0; k < lengthOfTable1; k++)
			{
				for (int p = 0; p < lengthOfTable2; p++)
				{
					if (table1ColumnList[k] == table2ColumnList[p])
					{
						newList1 = RowConvertedToList(table1Rows, columnNumber1[columnsIndex], k);
						newList2 = RowConvertedToList(table2Rows, columnNumber2[columnsIndex], p);
						compList = ComparedResult(newList1, newList2, columnsIndex, false);
						comparedData.Rows[tempForComparedData][newColumnSetup[columnsIndex]] = compList;
						columnsIndex++;



						newList1 = RowConvertedToList(table1Rows, columnNumber1[columnsIndex], k);
						newList2 = RowConvertedToList(table2Rows, columnNumber2[columnsIndex], p);
						compList = ComparedResult(newList1, newList2, columnsIndex, false);
						comparedData.Rows[tempForComparedData][newColumnSetup[columnsIndex]] = compList;
						columnsIndex++;



						newList1 = RowConvertedToList(table1Rows, columnNumber1[columnsIndex], k);
						newList2 = RowConvertedToList(table2Rows, columnNumber2[columnsIndex], p);
						compList = ComparedResult(newList1, newList2, columnNumber1[columnsIndex], true);
						comparedData.Rows[tempForComparedData][newColumnSetup[columnsIndex]] = compList;
						columnsIndex++;



						newList1 = RowConvertedToList(table1Rows, columnNumber1[columnsIndex], k);
						newList2 = RowConvertedToList(table2Rows, columnNumber2[columnsIndex], p);
						compList = ComparedResult(newList1, newList2, columnsIndex, false);
						comparedData.Rows[tempForComparedData][newColumnSetup[columnsIndex]] = compList;
						columnsIndex++;



						newList1 = RowConvertedToList(table1Rows, columnNumber1[columnsIndex], k);
						newList2 = RowConvertedToList(table2Rows, columnNumber2[columnsIndex], p);
						compList = ComparedResult(newList1, newList2,  columnsIndex, false);
						comparedData.Rows[tempForComparedData][newColumnSetup[columnsIndex]] = compList;
						columnsIndex++;

						newList1 = RowConvertedToList(table1Rows, columnNumber1[columnsIndex], k);
						newList2 = RowConvertedToList(table2Rows, columnNumber2[columnsIndex], p);
						compList = ComparedResult(newList1, newList2, columnsIndex, false);
						comparedData.Rows[tempForComparedData][newColumnSetup[columnsIndex]] = compList;
						columnsIndex++;

						newList1 = RowConvertedToList(table1Rows, columnNumber1[columnsIndex], k);
						newList2 = RowConvertedToList(table2Rows, columnNumber2[columnsIndex], p);
						compList = ComparedResult(newList1, newList2, columnsIndex, false);
						comparedData.Rows[tempForComparedData][newColumnSetup[columnsIndex]] = compList;


						tempForComparedData++;
						columnsIndex = 0;
					}
				}
			}
			return comparedData;
		}

		public List<string> ColumnConvertedToList(DataRow[] receivedDataRow, int column)
		{
			List<string> temporaryList = new List<string> { };
			var lengthOfRows = receivedDataRow.Count();
			for (int i = 0; i < lengthOfRows; i++)
			{
				temporaryList.Add(receivedDataRow[i][column].ToString());
			}
			return temporaryList;
		}

		private static List<string> RowConvertedToList(DataRow[] receivedDataRow, int columnValue, int rowValue)
		{
			string tempList = receivedDataRow[rowValue][columnValue].ToString();
			tempList = tempList.Replace(",", "");
			string[] namesArray = tempList.Split(' ');
			int lengthOfTable1 = namesArray.Length;
			List<string> newList = new List<string>(lengthOfTable1);
			newList.AddRange(namesArray);
			return newList;
		}

		private string ComparedResult(List<string> list1, List<string> list2, int currentColumn, bool isItDesignatorColumn)
		{
			int lengthOfList1 = list1.Count();
			int lengthOfList2 = list2.Count();
			int lengthOfWholeList = lengthOfList1 + lengthOfList2;
			String newString = "";
			var tempList1 = list1;
			List<string> newList = new List<string>();
			List<string> secondCopy = new List<string>();
			List<string> copyOfList1 = new List<string>();
			copyOfList1 = list1.ToList();
			var countCopyList1 = copyOfList1.Count();
			secondCopy = list1.ToList();
			int o = 0;
			for (int i = 0; i < lengthOfList1; i++)
			{
				o = 0;
				for (o = 0; o < lengthOfList2; o++)
				{
					if (list1[i].Any() != false && list2[o].Any() != false && list1[i].Contains(list2[o]))
					{
						list1.RemoveAt(i);
						list2.RemoveAt(o);
						lengthOfList1--;
						lengthOfList2--;
						o = lengthOfList2;
						i = -1;
					}
				}
			}

			if (string.Join("", list1) == string.Join("", list2))
			{
				list1.Clear();
				list2.Clear();
			}

			bool isEmpty1 = !list1.Any();
			bool isEmpty2 = !list2.Any();
			if (isEmpty1 == true && isEmpty2 == true )
			{

				for (int p = 0; p < countCopyList1; p++)
				{
					if (p != countCopyList1 - 1)
					{
						if (isItDesignatorColumn == true)
						{
							newString = newString + copyOfList1[p].ToString() + ", ";
						}
						else
						{
							newString = newString + copyOfList1[p].ToString() + " ";
						}
					}
					else
					{
						newString = newString + copyOfList1[p].ToString();
					}
				}
				return newString;

			}
			else
			{
				if (isEmpty1 == false || isEmpty2 == false)
				{
					list1.Insert(0, "\r\n");

				}
				newList = list2.Concat(list1).ToList();

			

				var newListCount = newList.Count();

				for (int i = 0; i < newList.Count(); i++)
				{
					if (i == lengthOfList2 - 1 || i == lengthOfList2)
					{
						newString = newString + newList[i].ToString();
					}
					else
					{
						if (i == newListCount - 1)
						{
							newString = newString + newList[i].ToString();
							
						}
						else
						{
							if (isItDesignatorColumn == true)
							{
								newString = newString + newList[i].ToString() + ", ";
							}
							else
							{
								newString = newString + newList[i].ToString() + " ";
							}
						}
					}
				}

				if ((isEmpty1 == false || isEmpty2 == false) && currentColumn != 0)
				{
					secondCopy = secondCopy.Except (list1).ToList();

					if (isItDesignatorColumn == true)
					{
						var countOfNewListnewList = newString.Length;
						string combinedString = string.Join(", ", secondCopy);
						newString = newString + "\r\n" + combinedString;
					}
					else
					{
						var countOfNewListnewList = newString.Length;
						string combinedString = string.Join(" ", secondCopy);
						newString = newString + "\r\n" + combinedString;
					}
				}

				return newString;
			}
		}
	}
}