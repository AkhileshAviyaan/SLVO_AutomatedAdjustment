using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;

public class Program()
{

	static void Main()
	{
		//Console.WriteLine("Give me Excel Path");
		//string excelPath=Console.ReadLine();

		//Console.WriteLine("Give me Checking Sheet Name");
		//string checkSheetName=Console.ReadLine();

		//Console.WriteLine("Give me Raw Sheet Name");
		//string rawSheetName=Console.ReadLine();

		string excelPath = "C:\\Users\\1akhi\\OneDrive\\Desktop\\SL\\Manohora\\SL14-Manohara Pul-ThulodharaOriginal5.xlsx";
		string checkSheetNameFirst = "SL14-Bridge Chyasal-South_%Cov";

		string rawSheetName = "SL14-Raw";
		string DirectionFirst = "SOUTH";
		string checkSheetNameSecond = "SL14-Bridge Chyasal-North_%Cov";
		string DirectionSecond = "NORTH";


		Adjust ad = new Adjust(excelPath, rawSheetName,checkSheetNameFirst,DirectionFirst,checkSheetNameSecond,DirectionSecond);
		ad.checkSolve();
	}

}