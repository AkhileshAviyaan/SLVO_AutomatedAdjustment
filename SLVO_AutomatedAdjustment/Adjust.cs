using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

internal class Adjust
{
	XLWorkbook workbook { get; set; }
	IXLWorksheet checkSheet { get; set; }
	IXLWorksheet rawSheet { get; set; }
	List<RawSelectedRows> rawSelectedRows { get; set; } = new List<RawSelectedRows>();
  List<RawSelectedRows> rawCarAndTaxi
	{
		get
		{
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().StartsWith("CAR")).ToList();
			if (filter != null)
			{
				return filter;
			}
			else
			{
				return new List<RawSelectedRows>();	
			}
		}
	}
	List<RawSelectedRows> rawTempo
	{
		get
		{
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().StartsWith("LARGE TEMPO") || n.VehicleType.Trim().StartsWith("ELECTRIC TEMPO")).ToList();
			if (filter != null)
			{
				return filter;
			}
			else
			{
				return new List<RawSelectedRows>();
			}
		}
	}
	List<RawSelectedRows> rawUtilityPickUp
	{
		get
		{
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().StartsWith("UTILITY")).ToList();
			if (filter != null)
			{
				return filter;
			}
			else
			{
				return new List<RawSelectedRows>();
			}
		}
	}
	List<RawSelectedRows> rawMicroBus
	{
		get
		{
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().StartsWith("MICRO")).ToList();
			if (filter != null)
			{
				return filter;
			}
			else
			{
				return new List<RawSelectedRows>();
			}
		}
	}
	List<RawSelectedRows> rawMiniBusDiff
	{
		get
		{
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().StartsWith("MINU") || n.VehicleType.Trim().StartsWith("MINIBUS") || n.VehicleType.Trim().StartsWith("BUS ELECTRIC")).ToList();
			if (filter != null)
			{
				return filter;
			}
			else
			{
				return new List<RawSelectedRows>();
			}
		}
	}
	List<RawSelectedRows> rawBigBusDiff
	{
		get
		{
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().StartsWith("BIG BUS")).ToList();
			if (filter != null)
			{
				return filter;
			}
			else
			{
				return new List<RawSelectedRows>();
			}
		}
	}
	List<RawSelectedRows> rawLightTruckDiff
	{
		get
		{
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().StartsWith("LIGHT TRUCK")).ToList();
			if (filter != null)
			{
				return filter;
			}
			else
			{
				return new List<RawSelectedRows>();
			}
		}
	}
	List<RawSelectedRows> rawHeavyTruckDiff
	{
		get
		{
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().StartsWith("HEAVY TRUCK")).ToList();
			if (filter != null)
			{
				return filter;
			}
			else
			{
				return new List<RawSelectedRows>();
			}
		}
	}
	List<RawSelectedRows> rawMultiAxleTruckDiff
	{
		get
		{
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().StartsWith("MULTI")).ToList();
			if (filter != null)
			{
				return filter;
			}
			else
			{
				return new List<RawSelectedRows>();
			}
		}
	}
	List<CheckSelectedRows> checkSelectedRows { get; set; } = new List<CheckSelectedRows>();
	int DataNoCheck { get; set; }
	int DataNoRow { get; set; }
	string Direction { get; set; }

	public Adjust(string excelPath, string checkSheetName, string rawSheetName, string direction)
	{
		workbook = new XLWorkbook(excelPath);
		checkSheet = workbook.Worksheet(checkSheetName);
		rawSheet = workbook.Worksheet(rawSheetName);
		Direction = direction;
	}
	public void checkSolve()
	{
		SaveToMemoryForProcess();
		Check();

	}

	void SaveToMemoryForProcess()
	{
		var rowsCheck = checkSheet.RowsUsed().ToList();
		var rowsRaw = rawSheet.RowsUsed().ToList();
		rowsCheck.RemoveRange(0, 3);
		rowsRaw.RemoveRange(0, 2);

		DataNoCheck = rowsCheck.Count();
		DataNoRow = rowsRaw.Count();

		foreach (var row in rowsCheck)
		{
			CheckSelectedRows csr = new CheckSelectedRows();
			int t, v;

			csr.ShortTime = row.Cell(3).GetString();

			t = row.Cell(9).GetValue<int>();
			v = row.Cell(8).GetValue<int>();
			csr.CarTaxiDiff = t - v;

			t = row.Cell(12).GetValue<int>();
			v = row.Cell(11).GetValue<int>();
			csr.TempoDiff = t - v;

			t = row.Cell(15).GetValue<int>();
			v = row.Cell(14).GetValue<int>();
			csr.UtilityPickUpDiff = t - v;

			t = row.Cell(18).GetValue<int>();
			v = row.Cell(17).GetValue<int>();
			csr.MicroBusDiff = t - v;

			t = row.Cell(21).GetValue<int>();
			v = row.Cell(20).GetValue<int>();
			csr.MiniBusDiff = t - v;

			t = row.Cell(24).GetValue<int>();
			v = row.Cell(23).GetValue<int>();
			csr.BigBusDiff = t - v;

			t = row.Cell(27).GetValue<int>();
			v = row.Cell(26).GetValue<int>();
			csr.LightTruckDiff = t - v;

			t = row.Cell(30).GetValue<int>();
			v = row.Cell(29).GetValue<int>();
			csr.HeavyTruckDiff = t - v;

			t = row.Cell(33).GetValue<int>();
			v = row.Cell(32).GetValue<int>();
			csr.MultiAxleTruckDiff = t - v;
			checkSelectedRows.Add(csr);

		}
		foreach (var row in rowsRaw)
		{
			RawSelectedRows rsr = new RawSelectedRows();
			var Direct = row.Cell(17).GetString();
			if (Direct.Trim().ToUpper() == Direction)
			{
				rsr.ShortTime = row.Cell(1).GetString();
				rsr.VehicleType = row.Cell(18).GetString();
				rsr.FullTime = row.Cell(20).GetString();
				rawSelectedRows.Add(rsr);
			}
		}
	}
	void Check()
	{
		//for (int i = 0; i < DataNoCheck; i++)
		//{
		//	var startTime = checkSelectedRows[i].ShortTime;
		//	var a = diff[i];
		//	if (a < 0)
		//	{
		//		if (i == 0)
		//		{
		//			if (diff[i + 1] <= 0)
		//			{

		//			}
		//			else
		//			{

		//			}

		//		}
		//		else if (i == DataNoCheck - 1)
		//		{

		//		}
		//		else
		//		{

		//		}

		//	}
		//}
	}
	class RawSelectedRows()
	{
		public string ShortTime { get; set; }
		public string VehicleType { get; set; }
		public string FullTime { get; set; }
	}
	class CheckSelectedRows()
	{
		public string ShortTime { get; set; }
		public double CarTaxiDiff { get; set; }
		public double TempoDiff { get; set; }
		public double UtilityPickUpDiff { get; set; }
		public double MicroBusDiff { get; set; }
		public double MiniBusDiff { get; set; }
		public double BigBusDiff { get; set; }
		public double LightTruckDiff { get; set; }
		public double HeavyTruckDiff { get; set; }
		public double MultiAxleTruckDiff { get; set; }
	}
}

