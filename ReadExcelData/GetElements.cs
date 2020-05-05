using System.Data;

namespace ReadExcelData
{
	class GetElements
	{
		public int[] GetColumnElements(DataTable receivedDataTable)
		{
			int[] columnSetup = new int[7];
			var amountOfColumns = receivedDataTable.Columns.Count;
			for (int i = 0; i < amountOfColumns; i++)
			{
				if (receivedDataTable.Rows[0][i].ToString() == "Quantity") columnSetup[0] = i;
				if (receivedDataTable.Rows[0][i].ToString() == "Part Number") columnSetup[1] = i;
				if (receivedDataTable.Rows[0][i].ToString() == "Designator") columnSetup[2] = i;
				if (receivedDataTable.Rows[0][i].ToString() == "Value") columnSetup[3] = i;
				if (receivedDataTable.Rows[0][i].ToString() == "Description") columnSetup[4] = i;
				if (receivedDataTable.Rows[0][i].ToString() == "Manufacturer") columnSetup[5] = i;
				if (receivedDataTable.Rows[0][i].ToString() == "Manufacturer Part Number") columnSetup[6] = i;
			}
			return columnSetup;
		}
	}
}
