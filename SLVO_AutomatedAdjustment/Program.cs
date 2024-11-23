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

		string excelPath = "C:\\Users\\1akhi\\OneDrive\\Desktop\\SL\\Manohora\\SL14-Manohara Pul-Thulodhara.xlsx";
		string checkSheetName = "SL14-Bridge Chyasal-North_%Cov";

		string rawSheetName = "SL14-Raw";
		string Direction = "NORTH";
		//string excelPath = "C:\\Users\\1akhi\\OneDrive\\Desktop\\SL\\Manohora\\S14 Manohara Pul Thulodhara.xlsx";
		//string checkSheetName = "SL14-Bridge Chyasal-South_%Cov";
		//string rawSheetName = "SL14-Raw";
		//string Direction = "SOUTH";


		Adjust ad = new Adjust(excelPath,checkSheetName,rawSheetName,Direction);
		ad.checkSolve();
	}

}