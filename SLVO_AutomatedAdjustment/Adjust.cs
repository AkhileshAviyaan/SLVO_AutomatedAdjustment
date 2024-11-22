using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

internal class Adjust
{
	XLWorkbook workbook { get; set; }
	IXLWorksheet checkSheet { get; set; }
	IXLWorksheet rawSheet { get; set; }
	List<RawSelectedRows> rawSelectedRows { get; set; } = new List<RawSelectedRows>();
	Dictionary<int, List<RawSelectedRows>> rawDict{get;set;}
	Dictionary<int, string> vehTypeDict = new Dictionary<int, string>() { { 9, "Taxi" }, { 12, "Tempo" }, { 15, "UtilityPickUp" }, { 18, "MicroBus" }, { 21, "MiniBus" }, { 24, "BigBus" }, { 27, "LightTruck" }, { 30, "HeavyTruck" }, { 33, "MultiAxleTruck" } };
	List<RawSelectedRows> rawCarAndTaxi
	{
		get
		{
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("CAR")).OrderBy(n => n.RowNo).ToList();
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
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("LARGE TEMPO") || n.VehicleType.Trim().ToUpper().StartsWith("ELECTRIC TEMPO")).OrderBy(n => n.RowNo).ToList();
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
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("UTILITY")).OrderBy(n => n.RowNo).ToList();
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
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("MICRO")).OrderBy(n => n.RowNo).ToList();
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
	List<RawSelectedRows> rawMiniBus
	{
		get
		{
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("MINU") || n.VehicleType.Trim().ToUpper().StartsWith("MINIBUS") || n.VehicleType.Trim().ToUpper().StartsWith("BUS ELECTRIC")).OrderBy(n => n.RowNo).ToList();
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
	List<RawSelectedRows> rawBigBus
	{
		get
		{
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("BIG BUS")).OrderBy(n => n.RowNo).ToList();
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
	List<RawSelectedRows> rawLightTruck
	{
		get
		{
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("LIGHT TRUCK")).OrderBy(n => n.RowNo).ToList();
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
	List<RawSelectedRows> rawHeavyTruck
	{
		get
		{
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("HEAVY TRUCK")).OrderBy(n => n.RowNo).ToList();
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
	List<RawSelectedRows> rawMultiAxleTruck
	{
		get
		{
			var filter = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("MULTI")).OrderBy(n => n.RowNo).ToList();
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
	List<CheckCell> checkCells { get; set; } = new List<CheckCell>();

	int DataNoCheck { get; set; }
	int DataNoRow { get; set; }
	string Direction { get; set; }

	public Adjust(string excelPath, string checkSheetName, string rawSheetName, string direction)
	{
		workbook = new XLWorkbook(excelPath);
		checkSheet = workbook.Worksheet(checkSheetName);
		rawSheet = workbook.Worksheet(rawSheetName);
		Direction = direction;
		rawDict = new Dictionary<int, List<RawSelectedRows>>() { { 9, rawCarAndTaxi }, { 12, rawTempo }, { 15, rawUtilityPickUp }, { 18, rawMicroBus }, { 21, rawMiniBus }, { 24, rawBigBus }, { 27, rawLightTruck }, { 30, rawHeavyTruck }, { 33, rawMultiAxleTruck } };
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
		rowsRaw.RemoveRange(0, 1);

		DataNoCheck = rowsCheck.Count();
		DataNoRow = rowsRaw.Count();

		foreach (var row in rowsCheck)
		{
			CheckSelectedRows csr = new CheckSelectedRows();
			csr.RowNo = row.RowNumber();
			csr.ShortTime = row.Cell(3).GetString();
			csr.Diff[9] = GetDiff(9,8,row);
			csr.Diff[12] = GetDiff(12, 11, row);
			csr.Diff[15] = GetDiff(15, 14, row);
			csr.Diff[18] = GetDiff(18, 17, row);
			csr.Diff[21] = GetDiff(21, 20, row);
			csr.Diff[24] = GetDiff(24, 23, row);
			csr.Diff[27] = GetDiff(27, 26, row);
			csr.Diff[30] = GetDiff(30, 29, row);
			csr.Diff[33] = GetDiff(33, 32, row);
			checkSelectedRows.Add(csr);
		}
		foreach (var row in rowsRaw)
		{
			RawSelectedRows rsr = new RawSelectedRows();
			var Direct = row.Cell(17).GetString();
			if (Direct.Trim().ToUpper() == Direction)
			{
				rsr.RowNo = row.RowNumber();
				rsr.ShortTime = row.Cell(1).GetString();
				rsr.VehicleType = row.Cell(18).GetString();
				rsr.FullTime = row.Cell(20).GetString();
				rawSelectedRows.Add(rsr);
			}
		}
	}
	int GetDiff(int t,int v, IXLRow row)
	{
		int t1 = row.Cell(t).GetValue<int>();
		int v1 = row.Cell(v).GetValue<int>();
		var diff = t1 - v1;
		if (diff < 0)
		{
			var c = new CheckCell();
			c.RowNo = row.RowNumber();
			c.ColNo = t;
			checkCells.Add(c);
		}
		else
		{
		}
		return t - v;
	}
	void Check()
	{
		for (int i = 0; i < checkCells.Count; i++)
		{
			var cellToCheck = checkCells[i];
			var cRow = cellToCheck.RowNo;
			var cCol=cellToCheck.ColNo;

			var csr = checkSelectedRows.Where(n => n.RowNo == cRow).FirstOrDefault();
			var shortTime = csr.ShortTime;
			int firstRowNo = 4;
			int lastRowNo = 4+checkSelectedRows.Count();
			int diff = csr.Diff[cCol];

			if (cRow == firstRowNo)
			{
				var csrBelow = checkSelectedRows.Where(n => n.RowNo == cRow+1).FirstOrDefault();
				if (csrBelow.Diff[cCol]<0)
				{
					for(int j = 0; j < Math.Abs(diff); j++)
					{
						var raws=rawDict[cCol];
						var rawsFiltedByTime = raws.Where(n => n.ShortTime == shortTime).OrderBy(n=>n.RowNo).ToList();
						rawsFiltedByTime.RemoveAt(j);
					}
				}
				else
				{
					var diffBelow = csrBelow.Diff[cCol];
					if (diffBelow >= Math.Abs(diff))
					{
						for (int j = Math.Abs(diff)-1; j >=0 ; j--)
						{
							var time=rawSheet.Cell(cRow, 20).GetString();

							var first = Convert.ToInt32(time.Substring(0, 11));
							var hr = Convert.ToInt32(time.Substring(11, 2));
							var min = Convert.ToInt32(time.Substring(14, 2));
							var sec = Convert.ToInt32(time.Substring(17, 2));
							Time t = new Time(hr, min, sec);
							t.Up();
							string newtime = first + hr + ":" + min + ":" + sec + ".000";
							rawSheet.Cell(cRow, 20).Value=newtime;
						}
					}
					else
					{
						var remDiff = Math.Abs(diff) - diffBelow;
						for (int j = diffBelow - 1; j >= 0; j--)
						{
							var time = rawSheet.Cell(cRow, 20).GetString();

							var first = Convert.ToInt32(time.Substring(0, 11));
							var hr = Convert.ToInt32(time.Substring(11, 2));
							var min = Convert.ToInt32(time.Substring(14, 2));
							var sec = Convert.ToInt32(time.Substring(17, 2));
							Time t = new Time(hr, min, sec);
							t.Up();
							string newtime = first + hr + ":" + min + ":" + sec + ".000";
							rawSheet.Cell(cRow, 20).Value = newtime;
						}
						for (int j = 0; j < Math.Abs(remDiff); j++)
						{
							var raws = rawDict[cCol];
							var rawsFiltedByTime = raws.Where(n => n.ShortTime == shortTime).OrderBy(n => n.RowNo).ToList();
							rawsFiltedByTime.RemoveAt(j);
						}
					}
				}
			}
			else if (cRow == lastRowNo)
			{
				var csrAbove = checkSelectedRows.Where(n => n.RowNo == cRow - 1).FirstOrDefault();
				if (csrAbove.Diff[cCol] < 0)
				{
					for (int j = 0; j < Math.Abs(diff); j++)
					{
						var raws = rawDict[cCol];
						var rawsFiltedByTime = raws.Where(n => n.ShortTime == shortTime).OrderBy(n => n.RowNo).ToList();
						rawsFiltedByTime.RemoveAt(j);
					}
				}
				else
				{
					var diffAbove = csrAbove.Diff[cCol];
					if (diffAbove >= Math.Abs(diff))
					{
						for (int j = 0; j <Math.Abs(diff); j++)
						{
							var time = rawSheet.Cell(cRow, 20).GetString();

							var first = Convert.ToInt32(time.Substring(0, 11));
							var hr = Convert.ToInt32(time.Substring(11, 2));
							var min = Convert.ToInt32(time.Substring(14, 2));
							var sec = Convert.ToInt32(time.Substring(17, 2));
							Time t = new Time(hr, min, sec);
							t.Down();
							string newtime = first + hr + ":" + min + ":" + sec + ".000";
							rawSheet.Cell(cRow, 20).Value = newtime;
						}
					}
					else
					{
						var remDiff = Math.Abs(diff) - diffAbove;
						for (int j =0; j < diffAbove; j--)
						{
							var time = rawSheet.Cell(cRow, 20).GetString();
							var first = Convert.ToInt32(time.Substring(0, 11));
							var hr = Convert.ToInt32(time.Substring(11, 2));
							var min = Convert.ToInt32(time.Substring(14, 2));
							var sec = Convert.ToInt32(time.Substring(17, 2));
							Time t = new Time(hr, min, sec);
							t.Down();
							string newtime = first + hr + ":" + min + ":" + sec + ".000";
							rawSheet.Cell(cRow, 20).Value = newtime;
						}
						for (int j = Math.Abs(remDiff)-1; j >=0 ; j--)
						{
							var raws = rawDict[cCol];
							var rawsFiltedByTime = raws.Where(n => n.ShortTime == shortTime).OrderBy(n => n.RowNo).ToList();
							rawsFiltedByTime.RemoveAt(j);
						}
					}
				}
			}
			else 
			{
				var csrAbove = checkSelectedRows.Where(n => n.RowNo == cRow - 1).FirstOrDefault();
				int remDiff=0;
				if (csrAbove.Diff[cCol] < 0)
				{
					for (int j = 0; j < Math.Abs(diff); j++)
					{
						var raws = rawDict[cCol];
						var rawsFiltedByTime = raws.Where(n => n.ShortTime == shortTime).OrderBy(n => n.RowNo).ToList();
						rawsFiltedByTime.RemoveAt(j);
					}
				}
				else
				{
					var diffAbove = csrAbove.Diff[cCol];
					if (diffAbove >= Math.Abs(diff))
					{
						for (int j = 0; j < Math.Abs(diff); j++)
						{
							var time = rawSheet.Cell(cRow, 20).GetString();

							var first = Convert.ToInt32(time.Substring(0, 11));
							var hr = Convert.ToInt32(time.Substring(11, 2));
							var min = Convert.ToInt32(time.Substring(14, 2));
							var sec = Convert.ToInt32(time.Substring(17, 2));
							Time t = new Time(hr, min, sec);
							t.Down();
							string newtime = first + hr + ":" + min + ":" + sec + ".000";
							rawSheet.Cell(cRow, 20).Value = newtime;
						}
					}
					else
					{
						 remDiff = Math.Abs(diff) - diffAbove;
						for (int j = 0; j < diffAbove; j--)
						{
							var time = rawSheet.Cell(cRow, 20).GetString();
							var first = Convert.ToInt32(time.Substring(0, 11));
							var hr = Convert.ToInt32(time.Substring(11, 2));
							var min = Convert.ToInt32(time.Substring(14, 2));
							var sec = Convert.ToInt32(time.Substring(17, 2));
							Time t = new Time(hr, min, sec);
							t.Down();
							string newtime = first + hr + ":" + min + ":" + sec + ".000";
							rawSheet.Cell(cRow, 20).Value = newtime;
						}
					}
				}

				

				var csrBelow = checkSelectedRows.Where(n => n.RowNo == cRow + 1).FirstOrDefault();
				if (remDiff < 0)
				{
					for (int j = 0; j < Math.Abs(diff); j++)
					{
						var raws = rawDict[cCol];
						var rawsFiltedByTime = raws.Where(n => n.ShortTime == shortTime).OrderBy(n => n.RowNo).ToList();
						rawsFiltedByTime.RemoveAt(j);
					}
				}
				else
				{
					var diffBelow = remDiff;
					if (diffBelow >= Math.Abs(diff))
					{
						for (int j = Math.Abs(diff) - 1; j >= 0; j--)
						{
							var time = rawSheet.Cell(cRow, 20).GetString();

							var first = Convert.ToInt32(time.Substring(0, 11));
							var hr = Convert.ToInt32(time.Substring(11, 2));
							var min = Convert.ToInt32(time.Substring(14, 2));
							var sec = Convert.ToInt32(time.Substring(17, 2));
							Time t = new Time(hr, min, sec);
							t.Up();
							string newtime = first + hr + ":" + min + ":" + sec + ".000";
							rawSheet.Cell(cRow, 20).Value = newtime;
						}
					}
					else
					{
					 remDiff = Math.Abs(diff) - diffBelow;
						for (int j = diffBelow - 1; j >= 0; j--)
						{
							var time = rawSheet.Cell(cRow, 20).GetString();

							var first = Convert.ToInt32(time.Substring(0, 11));
							var hr = Convert.ToInt32(time.Substring(11, 2));
							var min = Convert.ToInt32(time.Substring(14, 2));
							var sec = Convert.ToInt32(time.Substring(17, 2));
							Time t = new Time(hr, min, sec);
							t.Up();
							string newtime = first + hr + ":" + min + ":" + sec + ".000";
							rawSheet.Cell(cRow, 20).Value = newtime;
						}
						for (int j = 0; j < Math.Abs(remDiff); j++)
						{
							var raws = rawDict[cCol];
							var rawsFiltedByTime = raws.Where(n => n.ShortTime == shortTime).OrderBy(n => n.RowNo).ToList();
							rawsFiltedByTime.RemoveAt(j);
						}
					}
				}
			}

		}
	}
	class RawSelectedRows()
	{
		public int RowNo { get; set; }
		public string ShortTime { get; set; }
		public string VehicleType { get; set; }
		public string FullTime { get; set; }
	}
	class CheckSelectedRows()
	{
		public int RowNo { get; set; }
		public string ShortTime { get; set; }
		public Dictionary<int,int> Diff { get; set; }
	
	}
	class CheckCell
	{
		public int RowNo { get; set; }
		public int ColNo { get; set; }
	}
	class Time
	{
		public int Hr { get; set; }
		public int Min { get; set; }
		public int Sec { get; set; }
		public Time(int hr, int min, int sec)
		{
			this.Hr = hr;
			this.Min = min;
			this.Sec = sec;
		}

		public Time Up()
		{
			int remainder = Min % 15;
			if (remainder != 0)
			{
				Min += (15 - remainder);
			}

			// Handle overflow if Min becomes 60
			if (Min == 60)
			{
				Min = 0;
				Hr = (Hr + 1) % 24; // Ensure Hr stays within 24-hour format
			}

			Sec = 5; // Reset seconds
			return this;
		}

		public Time Down()
		{
			// Convert minutes to the previous 15-minute multiple
			int remainder = Min % 15;
			if (remainder != 0)
			{
				Min -= remainder;
				Min--;
			}

			Sec = 55; // Reset seconds
			return this;
		}
	}

}

