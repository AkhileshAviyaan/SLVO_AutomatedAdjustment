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

		string excelPath = "C:\\Users\\1akhi\\OneDrive\\Desktop\\SL\\SL13-Chyasal-Bagmati Bridge.xlsx";
		string checkSheetName = "SL3-North_%Cov";
		string rawSheetName = "SL13-Raw";
		string Direction = "NORTH";
		//string excelPath = "C:\\Users\\1akhi\\OneDrive\\Desktop\\SL\\Mid1.xlsm";
		//string checkSheetName = "SL6-Setopul-Maitidevi-West_%Cov";
		//string rawSheetName = "SL6-Raw";
		//string Direction = "WEST";


		Adjust ad = new Adjust(excelPath,checkSheetName,rawSheetName,Direction);
		ad.checkSolve();
	}

}