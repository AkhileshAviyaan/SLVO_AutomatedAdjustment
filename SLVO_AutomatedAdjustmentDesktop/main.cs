using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualBasic;
using System.Runtime.CompilerServices;

namespace SLVO_AutomatedAdjustmentDesktop
{
	public partial class Main : Form
	{
		public Main()
		{
			InitializeComponent();
		}
		public void Main_Load(object sender, EventArgs e)
		{

		}

		XLWorkbook Workbook { get; set; }
		string ExcelPath { get; set; }
		string rawSheetName { get; set; }
		string checkSheetNameFirst { get; set; }
		string checkSheetNameSecond { get; set; }
		string DirectionFirst { get; set; }
		string DirectionSecond { get; set; }
		List<string> WSLists { get; set; }
		void LoadExcel_Click(object sender, EventArgs e)
		{
				var dlg = new OpenFileDialog() { Title = "Find Excel Path", Filter = $"Excel File|*{".xlsx"}" };
				if (dlg.ShowDialog() == DialogResult.OK)
				{
					ExcelPath = dlg.FileName;
					txtExcelPath.Text = ExcelPath;
				}
			Cursor.Current = Cursors.WaitCursor;
			if (txtExcelPath.Text != null)
			{
				try
				{
					Workbook = new XLWorkbook(ExcelPath);
					WSLists = new List<string>();
					WSLists.Add("-Select Sheet-");
					foreach (var ws in Workbook.Worksheets)
					{
						WSLists.Add(ws.Name);
					}
					UpdateCombo();
					comboRawDataSheet.Items.AddRange(WSLists.ToArray());
					comboCheckSheetFirstDirection.Items.AddRange(WSLists.ToArray());
					comboCheckSheetSecondDirection.Items.AddRange(WSLists.ToArray());
					comboFirstDirection.Items.AddRange(new string[] { "-Select Direction-", "NORTH", "SOUTH", "EAST", "WEST" });
					comboSecondDirection.Items.AddRange(new string[] { "-Select Direction-", "NORTH", "SOUTH", "EAST", "WEST" });
					comboRawDataSheet.SelectedIndex = 0;
					comboCheckSheetFirstDirection.SelectedIndex = 0;
					comboCheckSheetSecondDirection.SelectedIndex = 0;
					comboFirstDirection.SelectedIndex = 0;
					comboSecondDirection.SelectedIndex = 0;
				}
				catch(Exception ex) 
				{
					MessageBox.Show(
						 "An error occurred while opening the Excel file. Please ensure that the file is not open in another program and try again.\n\nError Details: " + ex.Message,
						 "Error Opening File",
						 MessageBoxButtons.OK,
						 MessageBoxIcon.Error
					 );
				}

			}
			Cursor.Current = Cursors.Default;
		}
		void UpdateCombo() 
		{
			comboRawDataSheet.Items.Clear();
			comboCheckSheetFirstDirection.Items.Clear();
			comboCheckSheetSecondDirection.Items.Clear();
			comboFirstDirection.Items.Clear();
			comboSecondDirection.Items.Clear();
		}
		void Adjust_Click(object sender, EventArgs e)
		{
			bool result=false;
			Cursor.Current = Cursors.WaitCursor;
			if (UpdateParameters())
			{
			Adjust ad = new Adjust(ExcelPath, rawSheetName, checkSheetNameFirst, DirectionFirst, checkSheetNameSecond, DirectionSecond);
			result=ad.checkSolve();
			}
			Cursor.Current = Cursors.Default;
			if (result)
			{
				MessageBox.Show("Data Adjust in Excel file Successfully done", "Saving Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}

		}
		bool UpdateParameters()
		{
			rawSheetName = comboRawDataSheet.SelectedItem.ToString();
			checkSheetNameFirst = comboCheckSheetFirstDirection.SelectedItem.ToString();
			checkSheetNameSecond = comboCheckSheetSecondDirection.SelectedItem.ToString();
			DirectionFirst = comboFirstDirection.SelectedItem.ToString();
			DirectionSecond = comboSecondDirection.SelectedItem.ToString();
			if (rawSheetName == "-Select Sheet-" || checkSheetNameFirst == "-Select Sheet-" || checkSheetNameSecond == "-Select Sheet-")
			{
				MessageBox.Show("Please select a valid Sheet from dropdown!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return false;
			}
			if (DirectionFirst == "-Select Direction-" || DirectionSecond == "-Select Direction-")
			{
				MessageBox.Show("Please select a valid Direction from dropdown!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return false;
			}
			return true;
		}
	}
}
