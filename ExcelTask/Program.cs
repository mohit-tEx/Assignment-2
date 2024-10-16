using ClosedXML.Excel;

using var workbook = new XLWorkbook(@"C:\Book.xlsx");
var backupData = workbook.Worksheet("BackUP Data ");
var currentData = workbook.Worksheet("Current Data ");

int rowNum = 2;

Dictionary<string, string> dictionary = [];

while (backupData.Cell("A" + rowNum).Value.IsBlank == false)
{
	string userId = backupData.Cell("A" + rowNum).Value.GetText();
	string passwordHash = backupData.Cell("C" + rowNum).Value.GetText();

	dictionary.Add(userId, passwordHash);
	rowNum++;
}

rowNum = 2;
int numberOfChangedPassword = 0;

while (currentData.Cell("A" + rowNum).Value.IsBlank == false)
{
	string userId = currentData.Cell("A" + rowNum).Value.GetText();
	string currentPassword = currentData.Cell("C" + rowNum).Value.GetText();


	if (dictionary.ContainsKey(userId) && dictionary[userId].Equals(currentPassword) == false)
	{
		XLCellValue name = currentData.Cell("D" + rowNum).Value;

		Console.WriteLine(rowNum + " " + name);
		Console.WriteLine(dictionary[userId]);
		Console.WriteLine(currentPassword);
		Console.WriteLine("--------------------------------------------------");
		numberOfChangedPassword++;
	}

	rowNum++;
}
Console.WriteLine("Number of users with changed password: " + numberOfChangedPassword);