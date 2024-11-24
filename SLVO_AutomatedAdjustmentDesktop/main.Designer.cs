
namespace SLVO_AutomatedAdjustmentDesktop
{
	partial class Main
	{
		/// <summary>
		///  Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		///  Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		///  Required method for Designer support - do not modify
		///  the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			label1 = new Label();
			txtExcelPath = new TextBox();
			LoadExcel = new Button();
			label2 = new Label();
			label3 = new Label();
			label4 = new Label();
			label5 = new Label();
			label6 = new Label();
			comboRawDataSheet = new ComboBox();
			comboFirstDirection = new ComboBox();
			comboSecondDirection = new ComboBox();
			comboCheckSheetFirstDirection = new ComboBox();
			comboCheckSheetSecondDirection = new ComboBox();
			Adjust = new Button();
			SuspendLayout();
			// 
			// label1
			// 
			label1.AutoSize = true;
			label1.Location = new Point(217, 48);
			label1.Name = "label1";
			label1.Size = new Size(89, 25);
			label1.TabIndex = 0;
			label1.Text = "Excel Path";
			// 
			// txtExcelPath
			// 
			txtExcelPath.Location = new Point(348, 44);
			txtExcelPath.Name = "txtExcelPath";
			txtExcelPath.Size = new Size(784, 31);
			txtExcelPath.TabIndex = 1;
			// 
			// LoadExcel
			// 
			LoadExcel.Location = new Point(1152, 43);
			LoadExcel.Name = "LoadExcel";
			LoadExcel.Size = new Size(112, 34);
			LoadExcel.TabIndex = 2;
			LoadExcel.Text = "Load";
			LoadExcel.UseVisualStyleBackColor = true;
			LoadExcel.Click += LoadExcel_Click;
			// 
			// label2
			// 
			label2.AutoSize = true;
			label2.Location = new Point(170, 97);
			label2.Name = "label2";
			label2.Size = new Size(136, 25);
			label2.TabIndex = 3;
			label2.Text = "Raw Data Sheet";
			// 
			// label3
			// 
			label3.AutoSize = true;
			label3.Location = new Point(185, 145);
			label3.Name = "label3";
			label3.Size = new Size(121, 25);
			label3.TabIndex = 5;
			label3.Text = "First Direction";
			// 
			// label4
			// 
			label4.AutoSize = true;
			label4.Location = new Point(159, 194);
			label4.Name = "label4";
			label4.Size = new Size(147, 25);
			label4.TabIndex = 7;
			label4.Text = "Second Direction";
			// 
			// label5
			// 
			label5.AutoSize = true;
			label5.Location = new Point(79, 239);
			label5.Name = "label5";
			label5.Size = new Size(227, 25);
			label5.TabIndex = 9;
			label5.Text = "Check Sheet First Direction ";
			// 
			// label6
			// 
			label6.AutoSize = true;
			label6.Location = new Point(60, 291);
			label6.Name = "label6";
			label6.Size = new Size(246, 25);
			label6.TabIndex = 11;
			label6.Text = "Check Data Second Direction ";
			// 
			// comboRawDataSheet
			// 
			comboRawDataSheet.FormattingEnabled = true;
			comboRawDataSheet.Location = new Point(348, 94);
			comboRawDataSheet.Name = "comboRawDataSheet";
			comboRawDataSheet.Size = new Size(784, 33);
			comboRawDataSheet.TabIndex = 13;
			// 
			// comboFirstDirection
			// 
			comboFirstDirection.FormattingEnabled = true;
			comboFirstDirection.Location = new Point(348, 145);
			comboFirstDirection.Name = "comboFirstDirection";
			comboFirstDirection.Size = new Size(784, 33);
			comboFirstDirection.TabIndex = 14;
			// 
			// comboSecondDirection
			// 
			comboSecondDirection.FormattingEnabled = true;
			comboSecondDirection.Location = new Point(348, 191);
			comboSecondDirection.Name = "comboSecondDirection";
			comboSecondDirection.Size = new Size(784, 33);
			comboSecondDirection.TabIndex = 15;
			// 
			// comboCheckSheetFirstDirection
			// 
			comboCheckSheetFirstDirection.FormattingEnabled = true;
			comboCheckSheetFirstDirection.Location = new Point(348, 231);
			comboCheckSheetFirstDirection.Name = "comboCheckSheetFirstDirection";
			comboCheckSheetFirstDirection.Size = new Size(784, 33);
			comboCheckSheetFirstDirection.TabIndex = 16;
			// 
			// comboCheckSheetSecondDirection
			// 
			comboCheckSheetSecondDirection.FormattingEnabled = true;
			comboCheckSheetSecondDirection.Location = new Point(348, 288);
			comboCheckSheetSecondDirection.Name = "comboCheckSheetSecondDirection";
			comboCheckSheetSecondDirection.Size = new Size(784, 33);
			comboCheckSheetSecondDirection.TabIndex = 17;
			// 
			// Adjust
			// 
			Adjust.Location = new Point(601, 352);
			Adjust.Name = "Adjust";
			Adjust.Size = new Size(201, 34);
			Adjust.TabIndex = 18;
			Adjust.Text = "Adjust Data";
			Adjust.UseVisualStyleBackColor = true;
			Adjust.Click += this.Adjust_Click;
			// 
			// Main
			// 
			AutoScaleDimensions = new SizeF(10F, 25F);
			AutoScaleMode = AutoScaleMode.Font;
			ClientSize = new Size(1319, 450);
			Controls.Add(Adjust);
			Controls.Add(comboCheckSheetSecondDirection);
			Controls.Add(comboCheckSheetFirstDirection);
			Controls.Add(comboSecondDirection);
			Controls.Add(comboFirstDirection);
			Controls.Add(comboRawDataSheet);
			Controls.Add(label6);
			Controls.Add(label5);
			Controls.Add(label4);
			Controls.Add(label3);
			Controls.Add(label2);
			Controls.Add(LoadExcel);
			Controls.Add(txtExcelPath);
			Controls.Add(label1);
			Name = "Main";
			Text = "SLVO_Automated_Adjustment App";
			Load += Main_Load;
			ResumeLayout(false);
			PerformLayout();
		}

		#endregion

		private Label label1;
		private TextBox txtExcelPath;
		private Button LoadExcel;
		private Label label2;
		private Label label3;
		private Label label4;
		private Label label5;
		private Label label6;
		private ComboBox comboRawDataSheet;
		private ComboBox comboFirstDirection;
		private ComboBox comboSecondDirection;
		private ComboBox comboCheckSheetFirstDirection;
		private ComboBox comboCheckSheetSecondDirection;
		private Button Adjust;
	}
}
