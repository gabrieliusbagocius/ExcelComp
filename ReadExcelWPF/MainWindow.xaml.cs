using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.IO;
using System.Threading;
using System.Data;
using ClosedXML.Excel;
using System.Windows.Controls.Primitives;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Windows.Xps.Packaging;
using System.IO.Packaging;
using System.Reflection;
using System.Windows.Interop;

namespace ReadExcelWPF
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : System.Windows.Window
	{

		object SenderInherit { get; set; }
		DragEventArgs EInherit { get; set; }
		bool MousePosBool1 { get; set; }
		bool MousePosBool2 { get; set; }
		string CheckForDifference { get; set; }
		string FilePath1 { get; set; }
		string FilePath2 { get; set; }
		string FilePath3 { get; set; }
		double ScreenWidth = System.Windows.SystemParameters.MaximizedPrimaryScreenWidth;
		double ScreenHeight = System.Windows.SystemParameters.MaximizedPrimaryScreenHeight;
		bool FirstInit = true;
        string[] DataStringCopy { get; set; }
        System.Windows.Shapes.Rectangle RectangleCopy { get; set; }
		string CopyOfFilePath3 { get; set; }
        int CountRunTime { get; set; }
        int CountCopyRunTime { get; set; }
        int LastInt { get; set; }
        bool CheckWhichRan { get; set; }


        public MainWindow()
        {
            InitializeComponent();
            ProgramWindow.Height = Properties.Settings.Default.UserSavedHeight;
            ProgramWindow.MinHeight = Properties.Settings.Default.UserSavedHeight;
            MousePosBool1 = false;
            MousePosBool2 = false;
            FilePath1 = "";
            FilePath2 = "";
            FilePath3 = "";
            CopyOfFilePath3 = "";
            CountRunTime = 0;
            CountCopyRunTime = 0;
            LastInt = 1;
            CheckWhichRan = true;

        }




        public void ComparedButton_Click(object sender, EventArgs e)
		{
			OpenFileDialog openComparedFileDialog = new OpenFileDialog();
			openComparedFileDialog.Filter = "Excel Files(*.xls, *.xlsx) | *.xls; *.xlsx";
			openComparedFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
			openComparedFileDialog.FileName = "";

			if (openComparedFileDialog.ShowDialog() == true)
			{
				comparedTextBox.Text = openComparedFileDialog.FileName;
				FilePath1 = openComparedFileDialog.FileName;

				string acquiredFileName1 = System.IO.Path.GetFileName(openComparedFileDialog.FileName);
				secondComparedTextBox.Text = acquiredFileName1;
			}
		}

		public void ComparableButton_Click(object sender, EventArgs e)
		{
			OpenFileDialog openComparableFileDialog = new OpenFileDialog();
			openComparableFileDialog.Filter = "Excel Files(*.xls, *.xlsx) | *.xls; *.xlsx";
			openComparableFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
			openComparableFileDialog.FileName = "";

			if (openComparableFileDialog.ShowDialog() == true)
			{
				comparableTextBox.Text = openComparableFileDialog.FileName;
				FilePath2 = openComparableFileDialog.FileName;

				string acquiredFileName2 = System.IO.Path.GetFileName(openComparableFileDialog.FileName);
				secondComparableTextBox.Text = acquiredFileName2;
			}

		}

		public void ResultButton_Click(object sender, EventArgs e)
		{
            try
            {
                SaveFileDialog openResultFileDialog = new SaveFileDialog();
                openResultFileDialog.Filter = "Excel Files(*.xlsx) |*.xlsx";
				openResultFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                openResultFileDialog.InitialDirectory = @"c:\temp\";

                if (openResultFileDialog.ShowDialog() == true)
                {
                    resultTextBox.Text = openResultFileDialog.FileName;
                    string acquiredFilePath3 = openResultFileDialog.FileName;
                    FilePath3 = acquiredFilePath3;
                }
            }
            catch { }
		}

		public void ProgramStartButton_Click(object sender, EventArgs e)
		{
			ViewPDF.Source = null;
			ViewPDF.IsEnabled = false;
            bool doDelete = true;

            if (string.IsNullOrEmpty(FilePath3) == false)
            {
                doDelete = false;
                if (CountCopyRunTime == 0)
                {
                    CountRunTime = 0;
                    LastInt = 1;
                }
                CountCopyRunTime++;
            }
            ReadExcelData.BrainOfTheComparison startProgram = new ReadExcelData.BrainOfTheComparison();
			if (string.IsNullOrEmpty(FilePath1) == false && string.IsNullOrEmpty(FilePath2) == false)
			{
				if (string.IsNullOrEmpty(FilePath3) == true)
				{
					string path = FilePath1;
					string addSuffix1(string filePath, string suffix)
					{
						filePath = path;
						string fileDirectory = System.IO.Path.GetDirectoryName(path);
						string fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
						string fileExtension = System.IO.Path.GetExtension(filePath);
						return System.IO.Path.Combine(fileDirectory, String.Concat(fileName, suffix, fileExtension));
					}
                    
                    FilePath3 = addSuffix1(path, String.Format("({0})", "cmprd") + LastInt);
                    LastInt++;
					startProgram.Program(FilePath1,
					FilePath2, FilePath3);
				}
				else
				{

                    if (CountRunTime == 0)
                    {
                        CopyOfFilePath3 = FilePath3;
                        startProgram.Program(FilePath1,
                        FilePath2, FilePath3);
                    }
                    else
                    {
                        LastInt++;
                        string addSuffix2(string filePath, string suffix)
                        {
                            filePath = CopyOfFilePath3;
                            string fileDirectory = System.IO.Path.GetDirectoryName(CopyOfFilePath3);
                            string fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
                            string fileExtension = System.IO.Path.GetExtension(filePath);
                            return System.IO.Path.Combine(fileDirectory, String.Concat(fileName, suffix, fileExtension));
                        }
                        FilePath3 = addSuffix2(CopyOfFilePath3, String.Format("({0})", LastInt));
                        startProgram.Program(FilePath1,
                        FilePath2, FilePath3);

                    }
                }
                CountRunTime++;
				ProgramWindow.MinHeight = App.MyProgramMinHeight;
                ProgramWindow.MaxHeight = ScreenHeight;
                if (FirstInit == true) MinimizeOrMaximize_Click(null, null);
				LoadResultsPdfFile();
                if(doDelete == true) FilePath3 = "";
				StatusTextBox.Text = "The comparison has been completed";

			}
		}

		private void MinimizeOrMaximize_Click(object sender, RoutedEventArgs e)
		{
            try
            {
                if (FirstInit == true)
                {
                    Properties.Settings.Default.IsMaximized = !Properties.Settings.Default.IsMaximized;
                    FirstInit = false;
                }
                if (Properties.Settings.Default.IsMaximized == true)
                {
                    ProgramWindow.MinWidth = App.MyProgramMinWidth + App.MySecondGridBarWidth + App.MyBorderWidth - App.MyWidthDifferenceFix;
                    ProgramWindow.MaxWidth = ProgramWindow.MinWidth;
                    Maximize.Visibility = Visibility.Visible;
                    Minimize.Visibility = Visibility.Hidden;
                    Properties.Settings.Default.IsMaximized = false;

                }
                else
                {
                    ProgramWindow.MinWidth = App.MyProgramMinWidth + App.MySecondGridBarWidth + App.MyBorderWidth - App.MyWidthDifferenceFix;
                    Maximize.Visibility = Visibility.Hidden;
                    Minimize.Visibility = Visibility.Visible;
                    ProgramWindow.MaxWidth = ScreenWidth;
                    ProgramWindow.Width = Properties.Settings.Default.UserSavedWidth;
                    ProgramWindow.Height = Properties.Settings.Default.UserSavedHeight;

                    Properties.Settings.Default.IsMaximized = true;
                }
                Properties.Settings.Default.Save();
            }
            catch { }
		
		}

		private void Window_StateChanged(object sender, EventArgs e)
		{
			switch (this.WindowState)
			{
				case WindowState.Maximized:
					MinimizeMaximizeButton.IsHitTestVisible = false;
					Minimize.Visibility = Visibility.Hidden;
				    Maximize.Visibility = Visibility.Hidden;
					break;

				case WindowState.Minimized:
					break;

				case WindowState.Normal:
					MinimizeMaximizeButton.IsHitTestVisible = true;
					if (Properties.Settings.Default.IsMaximized == true)
					{
						Minimize.Visibility = Visibility.Visible;
					}
					else
					{
						Maximize.Visibility = Visibility.Visible;
					}
					break;
			}
		}

		public void LoadResultsPdfFile()
		{
			try
            {
                string pdfDocName = FilePath3.Replace(new FileInfo(FilePath3).Extension, "") + ".pdf";
                if (File.Exists(pdfDocName) == true)
                {
                    pdfDocName = (new DirectoryInfo(pdfDocName)).FullName;
                    ViewPDF.IsEnabled = true;
                    ViewPDF.Source = (new Uri(pdfDocName));
                    SecondGrid.Visibility = Visibility.Visible;
                    ProgramWindow.MaxWidth = ScreenWidth;
                    ProgramWindow.MinWidth = App.MyProgramMinWidth + App.MySecondGridBarWidth + App.MyBorderWidth;
                }
            }
            catch
            { }
		}

		private void ProgramWindow_SizeChanged(object sender, SizeChangedEventArgs e)
		{
			FrameworkElement pnlClient = this.Content as FrameworkElement;
			try
			{
				var currentWindowHeight = pnlClient.ActualHeight;
				var currentWindowWidth = pnlClient.ActualWidth;
				SecondGrid.Width = currentWindowWidth - App.MyProgramMinWidth;
				ViewPDF.Width = currentWindowWidth - FirstGrid.ActualWidth - SecondGridBar.ActualWidth - App.MyCompleteSecondGridWidths;

                if (currentWindowWidth > App.MyProgramMinWidth  && currentWindowWidth < ScreenWidth - App.MySaveByValue)
				{
					Properties.Settings.Default.UserSavedWidth = currentWindowWidth + App.MyCompleteSecondGridWidths;
					Properties.Settings.Default.Save();
				}
                if (currentWindowHeight < ScreenHeight - App.MySaveByValue && FirstInit == false)
                {
                    Properties.Settings.Default.UserSavedHeight = currentWindowHeight + App.MyHeightDifferenceFix;
                    App.MySavedProgramHeight = Properties.Settings.Default.UserSavedHeight;
                    Properties.Settings.Default.Save();
                }
            }
			catch { }
		}

		private void ComparedButton_MouseLeave(object sender, MouseEventArgs e)
		{
			comparedButton.IsHitTestVisible = false;
			rectangleCompared.Visibility = Visibility.Visible;
			rectangleComparable.Visibility = Visibility.Visible;
		}

		private void ComparableButton_MouseLeave(object sender, MouseEventArgs e)
		{
			comparableButton.IsHitTestVisible = false;
			rectangleCompared.Visibility = Visibility.Visible;
			rectangleComparable.Visibility = Visibility.Visible;
		}










		private void Whole_MouseMove(object sender, MouseEventArgs e)
		{

			System.Windows.Point pt = e.GetPosition(this);
			System.Windows.Point pointToWindow = Mouse.GetPosition(this);
			System.Windows.Shapes.Rectangle rectangle = sender as System.Windows.Shapes.Rectangle;

			if (rectangleCompared.IsMouseOver == true)
			{
				MousePosBool1 = true;
				Rectangle_Drop(SenderInherit, EInherit);
				comparedButton.IsHitTestVisible = true;
				rectangleCompared.Visibility = Visibility.Hidden;
			}

			if (rectangleComparable.IsMouseOver == true)
			{
				MousePosBool2 = true;
				Rectangle_Drop(SenderInherit, EInherit);
				comparableButton.IsHitTestVisible = true;
				rectangleComparable.Visibility = Visibility.Hidden;
			}
		}

		private Brush _previousFill = null;

		private void Rectangle_DragEnter(object sender, DragEventArgs e)
		{
			System.Windows.Shapes.Rectangle rectangle = sender as System.Windows.Shapes.Rectangle;
			if (rectangle != null)
			{
				_previousFill = rectangle.Fill;
				if (e.Data.GetDataPresent(DataFormats.FileDrop))
				{
					string[] dataString = (string[])e.Data.GetData(DataFormats.FileDrop);
					BrushConverter converter = new BrushConverter();
					rectangle.Fill = (Brush)converter.ConvertFromString("#BEE6FD");
					rectangle.Stroke = (Brush)converter.ConvertFromString("#FF2C628B");
				}
			}
		}

		private void Rectangle_DragOver(object sender, DragEventArgs e)
		{
			e.Effects = DragDropEffects.None;
			System.Windows.Shapes.Rectangle rectangle = sender as System.Windows.Shapes.Rectangle;
            RectangleCopy = rectangle;
			bool hitTestResult1 = HitTestResult.Equals(rectangle, rectangleCompared);
			bool hitTestResult2 = HitTestResult.Equals(rectangle, rectangleComparable);

			if (e.Data.GetDataPresent(DataFormats.FileDrop))
			{

				string[] dataString = (string[])e.Data.GetData(DataFormats.FileDrop);
                DataStringCopy = dataString;
				BrushConverter converter = new BrushConverter();
				rectangle.Fill = (Brush)converter.ConvertFromString("#BEE6FD");
				rectangle.Stroke = (Brush)converter.ConvertFromString("#FF2C628B");
			}

        }

        private void Rectangle_DragLeave(object sender, DragEventArgs e)
		{
			System.Windows.Shapes.Rectangle rectangle = sender as System.Windows.Shapes.Rectangle;
			if (rectangle != null)
			{
				BrushConverter converter = new BrushConverter();
				rectangle.Fill = (Brush)converter.ConvertFromString("#0054A6");
				rectangle.Stroke = (Brush)converter.ConvertFromString("#FF707070");
			}
		}

		private void Rectangle_Drop(object sender, DragEventArgs e)
		{
			SenderInherit = sender;
			EInherit = e;
			bool dropEnabled = true;
			System.Windows.Shapes.Rectangle rectangle = sender as System.Windows.Shapes.Rectangle;


            if (e != null)
			{

                if (e.Data.GetDataPresent(DataFormats.FileDrop))
				{
					string[] dataString = (string[])e.Data.GetData(DataFormats.FileDrop);
					if (System.IO.Path.GetExtension(dataString[0]).ToUpperInvariant() == ".XLSX" || System.IO.Path.GetExtension(dataString[0]).ToUpperInvariant() == ".XLS")
					{
						dropEnabled = true;
					}
					else
					{
						dropEnabled = false;
					}
					if (dropEnabled == true)
					{

						if (MousePosBool1 == true && MousePosBool2 == false && dataString[0] != null && CheckForDifference != dataString[0])
						{
							comparedTextBox.Text = dataString[0];
							string acquiredFileName1 = System.IO.Path.GetFileName(dataString[0]);
							secondComparedTextBox.Text = acquiredFileName1;

							FilePath1 = dataString[0];
							CheckForDifference = dataString[0];
						}


						if (MousePosBool1 == false && MousePosBool2 == true && dataString[0] != null && CheckForDifference != dataString[0])
						{
							comparableTextBox.Text = dataString[0];
							string acquiredFileName2 = System.IO.Path.GetFileName(dataString[0]);
							secondComparableTextBox.Text = acquiredFileName2;

							FilePath2 = dataString[0];
							CheckForDifference = dataString[0];
						}
					}
					BrushConverter converter = new BrushConverter();
					rectangle.Fill = (Brush)converter.ConvertFromString("#0054A6");
				}

        
			}
            if (e == null)
            {
                if (DataStringCopy != null)
                {
                    string[] dataString = DataStringCopy;
                    if (System.IO.Path.GetExtension(dataString[0]).ToUpperInvariant() != ".XLSX" || System.IO.Path.GetExtension(dataString[0]).ToUpperInvariant() != ".XLS")
                    {
                        dropEnabled = false;
                    }
                    else
                    {
                        dropEnabled = true;
                    }

                    if (dropEnabled == true)
                    {
                        if (MousePosBool1 == true && MousePosBool2 == false && dataString[0] != null && CheckForDifference != dataString[0])
                        {
                            comparedTextBox.Text = dataString[0];
                            string acquiredFileName1 = System.IO.Path.GetFileName(dataString[0]);
                            secondComparedTextBox.Text = acquiredFileName1;

                            FilePath1 = dataString[0];
                            CheckForDifference = dataString[0];
                        }
                        if (MousePosBool1 == false && MousePosBool2 == true && dataString[0] != null && CheckForDifference != dataString[0])
                        {
                            comparableTextBox.Text = dataString[0];
                            string acquiredFileName2 = System.IO.Path.GetFileName(dataString[0]);
                            secondComparableTextBox.Text = acquiredFileName2;

                            FilePath2 = dataString[0];
                            CheckForDifference = dataString[0];
                        }
                    }
                    BrushConverter converter = new BrushConverter();
                    RectangleCopy.Fill = (Brush)converter.ConvertFromString("#0054A6");
                }
            }
            MousePosBool1 = false;
			MousePosBool2 = false;
		}
	}
}
