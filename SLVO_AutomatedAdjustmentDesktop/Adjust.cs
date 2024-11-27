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
    Dictionary<int, string> vehTypeDict = new Dictionary<int, string>() { { 9, "Car/Taxi" }, { 12, "Electric Tempo" }, { 15, "Utility Pick Up" }, { 18, "Micro Bus (Hiace Type)" }, { 21, "Minubus (Regular)" }, { 24, "Big Bus" }, { 27, "Light Truck" }, { 30, "Heavy Truck" }, { 33, "Multi-axel Truck" } };
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
    public bool checkSolve()
    {
        SaveToMemoryForProcess();
        Check(true);
        Check(false);
        SaveWorkBook();
        return true;
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
                UpdateCsr(csr, 9, 8, row, true);
                UpdateCsr(csr, 12, 11, row, true);
                UpdateCsr(csr, 15, 14, row, true);
                UpdateCsr(csr, 18, 17, row, true);
                UpdateCsr(csr, 21, 20, row, true);
                UpdateCsr(csr, 24, 23, row, true);
                UpdateCsr(csr, 27, 26, row, true);
                UpdateCsr(csr, 30, 29, row, true);
                UpdateCsr(csr, 33, 32, row, true);
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
                UpdateCsr(csr, 9, 8, row, false);
                UpdateCsr(csr, 12, 11, row, false);
                UpdateCsr(csr, 15, 14, row, false);
                UpdateCsr(csr, 18, 17, row, false);
                UpdateCsr(csr, 21, 20, row, false);
                UpdateCsr(csr, 24, 23, row, false);
                UpdateCsr(csr, 27, 26, row, false);
                UpdateCsr(csr, 30, 29, row, false);
                UpdateCsr(csr, 33, 32, row, false);
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
    void UpdateCsr(CheckSelectedRows csr, int t, int v, IXLRow row, bool a)
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

        csr.Diff[t] = (int)diff;
        if (v1 == 0 & t1 != 0)
        {
            csr.ZeroCell.Add(t);
        }
        if (t1 != 0)
        {
            csr.Percentage[t] = (int)Math.Ceiling(v1 / t1 * 100);
            csr.Increment[t] = (int)(100 / t1);
        }
        else
        {
            csr.Percentage[t] = 100;
            csr.Increment[t] = 0;
        }
    }
    List<RawSelectedRows> rawsFiltedByTime { get; set; }
    List<RawSelectedRows> modifiedRaws { get; set; }
    void Updateraw()
    {
        foreach (var r in modifiedRaws)
        {
            rawsFiltedByTime.Remove(r);
        }
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
            var cellOfOneRow = cells.Where(n => n.RowNo == cRow).ToList();

            var csr = checkSelectedRows.Where(n => n.RowNo == cRow & n.IsFirstDirection == IsFirst).FirstOrDefault();


            var shortTime = csr.ShortTime;
            int firstRowNo = 5;
            int lastRowNo = firstRowNo + 64 - 1;
            int diff = csr.Diff[cCol];
            int AbsDiff = Math.Abs(diff);
            var raws = rawDict[cCol];
            rawsFiltedByTime = raws.Where(n => n.ShortTime == shortTime).OrderBy(n => n.RowNo).ToList();
            var FilteredCount = rawsFiltedByTime.Count();

            modifiedRaws = new List<RawSelectedRows>();

            //logic to adust within cell

            csr.ZeroCell = csr.ZeroCell.Where(n => n == 24 || n == 30).ToList();

            for (int j = csr.ZeroCell.Count() - 1; j >= 0; j--)
            {
                int a = csr.ZeroCell[j];
                if (csr.Diff[a] == 0)
                {
                    csr.ZeroCell.RemoveAt(j);
                }
            }
            int withOutDataCount = csr.ZeroCell.Count();

            if (cCol == 21)
            {
                if (csr.ZeroCell.Contains(24))
                {
                    if (FilteredCount > 0)
                    {
                        ModifyCellVehicleTypeChanged(rawsFiltedByTime[0], 0, 24);
                        modifiedRaws.Add(rawsFiltedByTime[0]);
                        csr.ZeroCell.RemoveAt(0);
                        csr.Diff[cCol]++;
                        csr.Diff[24]--;
                        AbsDiff--;
                    }
                }
            }
            else if (cCol == 27)
            {
                if (csr.ZeroCell.Contains(30))
                {
                    if (FilteredCount > 0)
                    {
                        ModifyCellVehicleTypeChanged(rawsFiltedByTime[0], 0, 30);
                        modifiedRaws.Add(rawsFiltedByTime[0]);
                        csr.ZeroCell.RemoveAt(0);
                        csr.Diff[cCol]++;
                        csr.Diff[30]--;
                        AbsDiff--;
                    }
                }
            }
            Updateraw();
            //int loopTo = 0;
            //if (AbsDiff == withOutDataCount)
            //{
            //	loopTo = withOutDataCount;
            //}
            //else if (withOutDataCount > AbsDiff)
            //{
            //	loopTo = AbsDiff;
            //}
            //else if (AbsDiff > withOutDataCount)
            //{
            //	loopTo = withOutDataCount;
            //}
            //for (int j = loopTo - 1; j >= 0; j--)
            //{
            //	if (FilteredCount>0)
            //	{
            //                 int col = csr.ZeroCell[j];
            //                 ModifyCellVehicleTypeChanged(rawsFiltedByTime[j], j, col);
            //                 modifiedRaws.Add(rawsFiltedByTime[j]);
            //                 csr.ZeroCell.RemoveAt(j);
            //                 csr.Diff[cCol]++;
            //                 csr.Diff[col]--;
            //             }
            //}
            //if (withOutDataCount >= AbsDiff)
            //{
            //	continue;
            //}
            diff = -(AbsDiff);

            if (FilteredCount == 0)
            {
                continue;
            }
            if (diff == 0)
            {
                continue;
            }
            FilteredCount = rawsFiltedByTime.Count();
            if (cRow == firstRowNo)
            {
                var csrBelow = checkSelectedRows.Where(n => n.RowNo == cRow + 1 & n.IsFirstDirection == IsFirst).FirstOrDefault();
                var diffBelow = csrBelow.Diff[cCol];
                if (diffBelow <= 0)
                {
                    for (int j = 0; j < Math.Abs(diff); j++)
                    {
                        Updateraw();
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
                            ModifyCellTimeShift(rawsFiltedByTime[j], j, "UP");
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
                            Updateraw();
                            ModifyCellTimeShift(rawsFiltedByTime[j], j, "UP");
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
                int startIndex = FilteredCount - 1;
                int endIndex = startIndex - Math.Abs(diff);
                if (endIndex < 0)
                {
                    endIndex = -1;
                }
                if (diffAbove <= 0)
                {
                    for (int j = startIndex; j > endIndex; j--)
                    {
                        DeleteCell(rawsFiltedByTime[j]);
                        csr.Diff[cCol]++;
                    }
                }
                else
                {
                    if (diffAbove >= Math.Abs(diff))
                    {
                        for (int j = 0; j < Math.Abs(diff); j++)
                        {
                            ModifyCellTimeShift(rawsFiltedByTime[j], j, "DOWN");
                            csrAbove.Diff[cCol]--;
                            csr.Diff[cCol]++;
                        }
                    }
                    else
                    {
                        var remDiff = Math.Abs(diff) - diffAbove;
                        for (int j = 0; j < diffAbove; j++)
                        {
                            ModifyCellTimeShift(rawsFiltedByTime[j], j, "DOWN");
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
                    continue;
                }
                else
                {
                    if (diffAbove >= Math.Abs(diff))
                    {
                        int endIndex = Math.Abs(diff);
                        for (int j = 0; j < endIndex; j++)
                        {
                            ModifyCellTimeShift(rawsFiltedByTime[j], j, "DOWN");
                            csrAbove.Diff[cCol]--;
                            csr.Diff[cCol]++;
                        }
                    }
                    else
                    {
                        var remDiff = Math.Abs(diff) - diffAbove;
                        for (int j = 0; j < diffAbove; j++)
                        {
                            ModifyCellTimeShift(rawsFiltedByTime[j], j, "DOWN");
                            csrAbove.Diff[cCol]--;
                            csr.Diff[cCol]++;
                        }
                        diff = -remDiff;
                        if (diffBelow <= 0)
                        {
                            for (int j = 0; j < Math.Abs(diff); j++)
                            {
                                Updateraw();
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
                                    ModifyCellTimeShift(rawsFiltedByTime[j], j, "UP");
                                    csrBelow.Diff[cCol]--;
                                    csr.Diff[cCol]++;
                                }
                            }
                            else
                            {
                                var remDiff1 = Math.Abs(diff) - diffBelow;
                                int startIndex = FilteredCount - 1;
                                int endIndex = startIndex - diffBelow;
                                if (endIndex < 0)
                                {
                                    continue;
                                }

                                for (int j = startIndex; j > endIndex; j--)
                                {
                                    ModifyCellTimeShift(rawsFiltedByTime[j], j, "UP");
                                    csrBelow.Diff[cCol]--;
                                    csr.Diff[cCol]++;
                                }
                                startIndex = startIndex - diffBelow;
                                endIndex = startIndex - remDiff1;
                                if (endIndex < 0)
                                {
                                    endIndex = -1;
                                }
                                if (startIndex >= 0)
                                {
                                    for (int j = startIndex; j > endIndex; j--)
                                    {
                                        DeleteCell(rawsFiltedByTime[j]);
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
    void ModifyCellTimeShift(RawSelectedRows rsr, int j, string Modify)
    {
        int rowNo = rsr.RowNo;
        bool CanBeModify = false;
        var time = rawSheet.Cell(rowNo, 20).GetString();

        var first = Convert.ToString(time.Substring(0, 11));

        var hr = Convert.ToInt32(time.Substring(11, 2));

        var min = Convert.ToInt32(time.Substring(14, 2));
        var sec = Convert.ToInt32(time.Substring(17, 2));
        Time t = new Time(hr, min, sec);
        if (Modify == "UP")
        {
            CanBeModify = t.UpPossible(j);
        }
        else if (Modify == "DOWN")
        {
            CanBeModify = t.DownPossible(j);
        }
        if (CanBeModify)
        {
            string newtime = first + t.Hr.ToString("D2") + ":" + t.Min.ToString("D2") + ":" + t.Sec.ToString("D2") + ".000";
            rawSheet.Cell(rowNo, 20).Value = newtime;
            modifiedRaws.Add(rsr);
        }
        else
        {
            var ts = first + t.Hr.ToString("D2") + ":" + t.Min.ToString("D2") + ":" + t.Sec.ToString("D2");
            DeleteCell(rsr);
            modifiedRaws.Add(rsr);
        }
    }
    void ModifyCellVehicleTypeChanged(RawSelectedRows rsr, int j, int cell)
    {
        int rowNo = rsr.RowNo;
        rawSheet.Cell(rowNo, 18).Value = vehTypeDict[cell];
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
        public List<int> ZeroCell { get; set; } = new List<int>();
        public Dictionary<int, int> Diff { get; set; } = new Dictionary<int, int>();
        public Dictionary<int, int> Percentage { get; set; } = new Dictionary<int, int>();
        public Dictionary<int, int> Increment { get; set; } = new Dictionary<int, int>();
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
        public bool UpPossible(int j)
        {
            int remainder = Min % 15;
            if (remainder > 10)
            {
                if (Min >= 45 & Min < 60)
                {
                    Min = 1;
                    Hr = Hr + 1;
                }
                else if (Min >= 30 & Min < 45)
                {
                    Min = 46;
                }
                else if (Min >= 15 & Min < 30)
                {
                    Min = 31;
                }
                else if (Min >= 0 & Min < 15)
                {
                    Min = 16;
                }
                Random rdm = new Random();
                Sec = 5 + rdm.Next(0, 30);
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool DownPossible(int j)
        {
            // Convert minutes to the previous 15-minute multiple
            int remainder = Min % 15;
            if (remainder < 5)
            {
                if (Min >= 45 & Min < 60)
                {
                    Min = 44;
                }
                else if (Min >= 30 & Min < 45)
                {
                    Min = 29;
                }
                else if (Min >= 15 & Min < 30)
                {
                    Min = 14;
                }
                else if (Min >= 0 & Min < 15)
                {
                    Min = 59;
                    Hr = (Hr - 1);
                }
                Random rdm = new Random();
                Sec = 5 + rdm.Next(0, 30);
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}

