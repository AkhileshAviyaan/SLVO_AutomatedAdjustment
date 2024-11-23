using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.ConstrainedExecution;
using System.Runtime.Intrinsics.Arm;
using System.Text;
using System.Threading.Tasks;

internal class Adjust
{
	XLWorkbook workbook { get; set; }
	IXLWorksheet checkSheetFirst { get; set; }
	IXLWorksheet checkSheetSecond { get; set; }
	IXLWorksheet rawSheet { get; set; }
	string Path { get; set; }
	List<RawSelectedRows> rawSelectedRowsFirst { get; set; } = new List<RawSelectedRows>();
	List<RawSelectedRows> rawSelectedRowsSecond { get; set; } = new List<RawSelectedRows>();
	Dictionary<int, List<RawSelectedRows>> rawDict { get; set; }
	Dictionary<int, string> vehTypeDict = new Dictionary<int, string>() { { 9, "Taxi" }, { 12, "Tempo" }, { 15, "UtilityPickUp" }, { 18, "MicroBus" }, { 21, "MiniBus" }, { 24, "BigBus" }, { 27, "LightTruck" }, { 30, "HeavyTruck" }, { 33, "MultiAxleTruck" } };
	void SaveWorkBook()
	{
		this.workbook.SaveAs(this.Path);
	}

	List<RawSelectedRows> rawCarAndTaxi;
	List<RawSelectedRows> rawTempo;
	List<RawSelectedRows> rawUtilityPickUp;
	List<RawSelectedRows> rawMicroBus;
	List<RawSelectedRows> rawMiniBus;
	List<RawSelectedRows> rawBigBus;
	List<RawSelectedRows> rawLightTruck;
	List<RawSelectedRows> rawHeavyTruck;
	List<RawSelectedRows> rawMultiAxleTruck;
	void UpdateVehicleList(List<RawSelectedRows> rawSelectedRows)
	{
		rawCarAndTaxi = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("CAR")).OrderBy(n => n.RowNo).ToList();
		rawTempo = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("LARGE TEMPO") || n.VehicleType.Trim().ToUpper().StartsWith("ELECTRIC TEMPO")).OrderBy(n => n.RowNo).ToList();

		rawUtilityPickUp = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("UTILITY")).OrderBy(n => n.RowNo).ToList();

		rawMicroBus = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("MICRO")).OrderBy(n => n.RowNo).ToList();

		rawMiniBus = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("MINU") || n.VehicleType.Trim().ToUpper().StartsWith("MINIBUS") || n.VehicleType.Trim().ToUpper().StartsWith("BUS ELECTRIC")).OrderBy(n => n.RowNo).ToList();

		rawBigBus = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("BIG BUS")).OrderBy(n => n.RowNo).ToList();

		rawLightTruck = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("LIGHT TRUCK")).OrderBy(n => n.RowNo).ToList();

		rawHeavyTruck = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("HEAVY TRUCK")).OrderBy(n => n.RowNo).ToList();

		rawMultiAxleTruck = rawSelectedRows.Where(n => n.VehicleType.Trim().ToUpper().StartsWith("MULTI")).OrderBy(n => n.RowNo).ToList();
	}

	List<CheckSelectedRows> checkSelectedRows { get; set; } = new List<CheckSelectedRows>();
	List<CheckCell> checkCells { get; set; } = new List<CheckCell>();

	int DataNoCheck { get; set; }
	int DataNoRow { get; set; }
	string DirectionFirst { get; set; }
	string DirectionSecond { get; set; }

	public Adjust(string excelPath, string rawSheetName, string checkSheetNameFirst, string directionFirst, string checkSheetNameSecond, string directionSecond)
	{
		this.Path = excelPath;
		workbook = new XLWorkbook(excelPath);
		rawSheet = workbook.Worksheet(rawSheetName);
		checkSheetFirst = workbook.Worksheet(checkSheetNameFirst);
		checkSheetSecond = workbook.Worksheet(checkSheetNameSecond);
		DirectionFirst = directionFirst;
		DirectionSecond = directionSecond;
	}
	public void checkSolve()
	{
		SaveToMemoryForProcess();
		Check(true);
		Check(false);
		SaveWorkBook();
	}
	IXLRows rowsCheckFirst;
	IXLRows rowsCheckSecond;
	IXLRows rowsRaw;
	void SaveToMemoryForProcess()
	{
		rowsCheckFirst = checkSheetFirst.RowsUsed();
		rowsCheckSecond = checkSheetSecond.RowsUsed();
		rowsRaw = rawSheet.RowsUsed();

		foreach (var row in rowsCheckFirst)
		{
			CheckSelectedRows csr = new CheckSelectedRows();
			csr.RowNo = row.RowNumber();
			if (csr.RowNo > 4)
			{
				csr.ShortTime = row.Cell(3).GetString();
				csr.IsFirstDirection = true;
				csr.Diff[9] = GetDiff(9, 8, row, true);
				csr.Diff[12] = GetDiff(12, 11, row, true);
				csr.Diff[15] = GetDiff(15, 14, row, true);
				csr.Diff[18] = GetDiff(18, 17, row, true);
				csr.Diff[21] = GetDiff(21, 20, row, true);
				csr.Diff[24] = GetDiff(24, 23, row, true);
				csr.Diff[27] = GetDiff(27, 26, row, true);
				csr.Diff[30] = GetDiff(30, 29, row, true);
				csr.Diff[33] = GetDiff(33, 32, row, true);
				checkSelectedRows.Add(csr);
			}
		}
		foreach (var row in rowsCheckSecond)
		{
			CheckSelectedRows csr = new CheckSelectedRows();
			csr.RowNo = row.RowNumber();
			if (csr.RowNo > 4)
			{
				csr.ShortTime = row.Cell(3).GetString();
				csr.IsFirstDirection = false;
				csr.Diff[9] = GetDiff(9, 8, row, false);
				csr.Diff[12] = GetDiff(12, 11, row, false);
				csr.Diff[15] = GetDiff(15, 14, row, false);
				csr.Diff[18] = GetDiff(18, 17, row, false);
				csr.Diff[21] = GetDiff(21, 20, row, false);
				csr.Diff[24] = GetDiff(24, 23, row, false);
				csr.Diff[27] = GetDiff(27, 26, row, false);
				csr.Diff[30] = GetDiff(30, 29, row, false);
				csr.Diff[33] = GetDiff(33, 32, row, false);
				checkSelectedRows.Add(csr);
			}
		}
		foreach (var row in rowsRaw)
		{
			RawSelectedRows rsr = new RawSelectedRows();
			var Direct = row.Cell(17).GetString();
			if (Direct.Trim().ToUpper() == DirectionFirst)
			{
				rsr.RowNo = row.RowNumber();
				rsr.ShortTime = row.Cell(1).GetString();
				rsr.VehicleType = row.Cell(18).GetString();
				rsr.FullTime = row.Cell(20).GetString();
				rawSelectedRowsFirst.Add(rsr);
			}
			else if (Direct.Trim().ToUpper() == DirectionSecond)
			{
				rsr.RowNo = row.RowNumber();
				rsr.ShortTime = row.Cell(1).GetString();
				rsr.VehicleType = row.Cell(18).GetString();
				rsr.FullTime = row.Cell(20).GetString();
				rawSelectedRowsSecond.Add(rsr);
			}
		}
	}
	int GetDiff(int t, int v, IXLRow row, bool a)
	{
		double t1 = row.Cell(t).GetDouble();
		double v1 = row.Cell(v).GetDouble();
		double diff = t1 - v1;

		if (diff < 0)
		{
			var c = new CheckCell
			{
				RowNo = row.RowNumber(),
				ColNo = t,
				IsFirstDirection = a
			};
			checkCells.Add(c);
		}

		return (int)diff;
	}

	void Check(bool IsFirst)
	{
		if (IsFirst)
		{
			UpdateVehicleList(rawSelectedRowsFirst);
		}
		else
		{
			UpdateVehicleList(rawSelectedRowsSecond);
		}
		rawDict = new Dictionary<int, List<RawSelectedRows>>() { { 9, rawCarAndTaxi }, { 12, rawTempo }, { 15, rawUtilityPickUp }, { 18, rawMicroBus }, { 21, rawMiniBus }, { 24, rawBigBus }, { 27, rawLightTruck }, { 30, rawHeavyTruck }, { 33, rawMultiAxleTruck } };

		var cells = checkCells.Where(n => n.IsFirstDirection == IsFirst).ToList();
		for (int i = 0; i < cells.Count; i++)
		{
			var cellToCheck = cells[i];
			var cRow = cellToCheck.RowNo;
			var cCol = cellToCheck.ColNo;

			var csr = checkSelectedRows.Where(n => n.RowNo == cRow & n.IsFirstDirection == IsFirst).FirstOrDefault();
			var shortTime = csr.ShortTime;
			int firstRowNo = 5;
			int lastRowNo = firstRowNo + 64 - 1;
			int diff = csr.Diff[cCol];

			var raws = rawDict[cCol];
			var rawsFiltedByTime = raws.Where(n => n.ShortTime == shortTime).OrderBy(n => n.RowNo).ToList();
			var FilteredCount = rawsFiltedByTime.Count();
			if (FilteredCount == 0)
			{
				continue;
			}
			if (cRow == firstRowNo)
			{
				var csrBelow = checkSelectedRows.Where(n => n.RowNo == cRow + 1 & n.IsFirstDirection == IsFirst).FirstOrDefault();
				var diffBelow = csrBelow.Diff[cCol];
				if (diffBelow <= 0)
				{
					for (int j = 0; j < Math.Abs(diff); j++)
					{
						DeleteCell(rawsFiltedByTime[j]);
						csr.Diff[cCol]++;
					}
				}
				else
				{
					if (diffBelow >= Math.Abs(diff))
					{
						int startIndex = FilteredCount - 1;
						int endIndex = startIndex + diff;
						for (int j = startIndex; j > endIndex; j--)
						{
							ModifyCell(rawsFiltedByTime[j], j, "UP");
							csrBelow.Diff[cCol]--;
							csr.Diff[cCol]++;
						}
					}
					else
					{
						var remDiff = Math.Abs(diff) - diffBelow;
						int startIndex = FilteredCount - 1;
						int endIndex = startIndex - remDiff;
						for (int j = startIndex; j > endIndex; j--)
						{
							DeleteCell(rawsFiltedByTime[j]);
							csr.Diff[cCol]++;
						}
						startIndex = startIndex - remDiff;
						endIndex = startIndex - diffBelow;
						for (int j = startIndex; j > endIndex; j--)
						{
							ModifyCell(rawsFiltedByTime[j], j, "UP");
							csrBelow.Diff[cCol]--;
							csr.Diff[cCol]++;
						}
					}
				}
			}
			else if (cRow == lastRowNo)
			{
				var csrAbove = checkSelectedRows.Where(n => n.RowNo == cRow - 1 & n.IsFirstDirection == IsFirst).FirstOrDefault();
				var diffAbove = csrAbove.Diff[cCol];
				if (diffAbove <= 0)
				{
					for (int j = FilteredCount - 1; j > Math.Abs(diff); j--)
					{
						DeleteCell(rawsFiltedByTime[j]);
						csr.Diff[cCol]++;
					}
				}
				else
				{
					if (diffAbove >= Math.Abs(diff))
					{
						int endIndex = Math.Abs(diff);
						for (int j = 0; j < endIndex; j++)
						{
							ModifyCell(rawsFiltedByTime[j], j, "DOWN");
							csrAbove.Diff[cCol]--;
							csr.Diff[cCol]++;
						}
					}
					else
					{
						var remDiff = Math.Abs(diff) - diffAbove;
						for (int j = 0; j < diffAbove; j++)
						{
							ModifyCell(rawsFiltedByTime[j], j, "DOWN");
							csrAbove.Diff[cCol]--;
							csr.Diff[cCol]++;
						}
						for (int j = diffAbove; j < Math.Abs(diff); j++)
						{
							DeleteCell(rawsFiltedByTime[j]);
							csr.Diff[cCol]++;
						}
					}
				}
			}
			else
			{
				var csrAbove = checkSelectedRows.Where(n => n.RowNo == cRow - 1 & n.IsFirstDirection == IsFirst).FirstOrDefault();
				var csrBelow = checkSelectedRows.Where(n => n.RowNo == cRow + 1 & n.IsFirstDirection == IsFirst).FirstOrDefault();
				var diffAbove = csrAbove.Diff[cCol];
				var diffBelow = csrBelow.Diff[cCol];
				if (diffAbove <= 0 & diffBelow <= 0)
				{
					if (FilteredCount == Math.Abs(diff))
					{
						for (int j = 0; j < Math.Abs(diff); j++)
						{
							DeleteCell(rawsFiltedByTime[j]);
							csr.Diff[cCol]++;
						}
					}
					else if (FilteredCount > Math.Abs(diff))
					{
						for (int j = 0; j < Math.Abs(diff); j++)
						{
							DeleteCell(rawsFiltedByTime[j]);
							csr.Diff[cCol]++;
						}
					}
					else if (FilteredCount < Math.Abs(diff))
					{
						for (int j = 0; j < FilteredCount; j++)
						{
							DeleteCell(rawsFiltedByTime[j]);
							csr.Diff[cCol]++;
						}
					}
				}
				else
				{
					if (diffAbove >= Math.Abs(diff))
					{
						int endIndex = Math.Abs(diff);
						for (int j = 0; j < endIndex; j++)
						{
							ModifyCell(rawsFiltedByTime[j], j, "DOWN");
							csrAbove.Diff[cCol]--;
							csr.Diff[cCol]++;
						}
					}
					else
					{
						var remDiff = Math.Abs(diff) - diffAbove;
						List<RawSelectedRows> modifiedRaws = new List<RawSelectedRows>();
						for (int j = 0; j < diffAbove; j++)
						{
							ModifyCell(rawsFiltedByTime[j], j, "DOWN");
							modifiedRaws.Add(rawsFiltedByTime[j]);
							csrAbove.Diff[cCol]--;
							csr.Diff[cCol]++;
						}
						diff = -remDiff;
						foreach (var r in modifiedRaws)
						{
							rawsFiltedByTime.Remove(r);
						}
						if (diffBelow <= 0)
						{
							for (int j = 0; j < Math.Abs(diff); j++)
							{
								DeleteCell(rawsFiltedByTime[j]);
								csr.Diff[cCol]++;
							}
						}
						else
						{
							if (diffBelow >= Math.Abs(diff))
							{
								int startIndex = FilteredCount - 1 - modifiedRaws.Count();
								int endIndex = startIndex + diff;
								for (int j = startIndex; j > endIndex; j--)
								{
									ModifyCell(rawsFiltedByTime[j], j, "UP");
									csrBelow.Diff[cCol]--;
									csr.Diff[cCol]++;
								}
							}
							else
							{
								var remDiff1 = Math.Abs(diff) - diffBelow;
								int startIndex = FilteredCount - 1;
								int endIndex = startIndex - remDiff1;
								if (endIndex < 0)
								{
									continue;
								}
								for (int j = startIndex; j > endIndex; j--)
								{
									DeleteCell(rawsFiltedByTime[j]);
									csr.Diff[cCol]++;
								}
								startIndex = startIndex - remDiff1;
								endIndex = startIndex - diffBelow;
								if (startIndex > 0)
								{
									for (int j = startIndex; j > endIndex; j--)
									{
										ModifyCell(rawsFiltedByTime[j], j, "UP");
										csrBelow.Diff[cCol]--;
										csr.Diff[cCol]++;
									}
								}
							}
						}
					}
				}
			}

		}
	}
	void DeleteCell(RawSelectedRows rsr)
	{
		int rowNo = rsr.RowNo;
		for (int k = 1; k <= 25; k++)
		{
			rawSheet.Cell(rowNo, k).Clear();
		}
	}
	string UpTime(string a)
	{
		string[] parts = a.Split(':');
		int hr = Convert.ToInt32(parts[0]);
		int min = Convert.ToInt32(parts[1]);
		double time = hr / 24f + min / 60f / 24f + 15 / 60f / 24f;
		int hr2 = (int)(time * 24);
		int min2 = (int)Math.Ceiling((time - hr2 / 24f) * 24 * 60);
		string b = hr2 + ":" + min2.ToString("D2") + ":00";
		return b;
	}
	string DownTime(string a)
	{
		string[] parts = a.Split(':');
		int hr = Convert.ToInt32(parts[0]);
		int min = Convert.ToInt32(parts[1]);
		double time = hr / 24f + min / 60f / 24f - 15 / 60f / 24f;
		int hr2 = (int)(time * 24);
		int min2 = (int)Math.Ceiling((time - hr2 / 24f) * 24 * 60);
		string b = hr2 + ":" + min2.ToString("D2") + ":00";
		return b;
	}
	void ModifyCell(RawSelectedRows rsr, int j, string Modify)
	{
		int rowNo = rsr.RowNo;
		var time = rawSheet.Cell(rowNo, 20).GetString();
		var first = Convert.ToString(time.Substring(0, 11));
		var hr = Convert.ToInt32(time.Substring(11, 2));
		var min = Convert.ToInt32(time.Substring(14, 2));
		var sec = Convert.ToInt32(time.Substring(17, 2));
		Time t = new Time(hr, min, sec);
		if (Modify == "UP")
		{
			t.Up(j);
		}
		else if (Modify == "DOWN")
		{
			t.Down(j);
		}
		string newtime = first + t.Hr.ToString("D2") + ":" + t.Min.ToString("D2") + ":" + t.Sec.ToString("D2") + ".000";
		rawSheet.Cell(rowNo, 20).Value = newtime;
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
		public bool IsFirstDirection { get; set; }
		public int RowNo { get; set; }
		public string ShortTime { get; set; }
		public Dictionary<int, int> Diff { get; set; } = new Dictionary<int, int>();

	}
	class CheckCell
	{
		public bool IsFirstDirection { get; set; }
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

		public Time Up(int j)
		{
			int remainder = Min % 15;


			if (remainder == 0 & Min == 45)
			{
				Min = 1;
				Hr++;
			}
			else if (remainder == 0 & Min == 30)
			{
				Min += 15;
			}
			else if (remainder == 0 & Min == 15)
			{
				Min += 15;
			}
			else if (remainder == 0 & Min == 0)
			{
				Min += 15;
			}
			else if (remainder != 0)
			{
				Min += (15 - remainder);
			}
			if (Min == 60)
			{
				Min = 1;
				Hr = (Hr + 1) % 24;
			}

			Random rdm = new Random();
			Sec = 5 + rdm.Next(0, 30);  // Reset seconds
			return this;
		}

		public Time Down(int j)
		{
			// Convert minutes to the previous 15-minute multiple
			int remainder = Min % 15;
			if (remainder == 0 & Min == 45)
			{
				Min -= 15;
			}
			else if (remainder == 0 & Min == 30)
			{
				Min -= 15;
			}
			else if (remainder == 0 & Min == 15)
			{
				Min -= 15;
			}
			else if (remainder == 0 & Min == 0)
			{
				Min = 59;
				Hr = (Hr - 1) % 24;
			}
			else if (remainder != 0 & remainder < 15)
			{
				Min -= remainder;
				Min--;
			}
			else if (remainder != 0)
			{
				Min -= remainder;
				Min--;
			}
			if (Min == -1)
			{
				Min = 59;
				Hr = (Hr - 1) % 24;
			}
			Random rdm = new Random();
			Sec = 5 + rdm.Next(0, 30); // Reset seconds
			return this;
		}
	}

}

