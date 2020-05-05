using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using DocumentFormat.OpenXml.Drawing.Diagrams;

namespace ReadExcelWPF
{
	/// <summary>
	/// Interaction logic for App.xaml
	/// </summary>
	public partial class App : Application
	{
		//DefaultProgramWidthOrHeight//
		public static int MyProgramMinHeight { get; set; }
		public static int MyProgramMinWidth { get; set; }
		public static int MyProgramMinimizedWidth { get; set; }
		public static int MyFirstGridWidth { get; set; }
		public static int MySecondGridWidth { get; set; }
        public static int MySecondGridBarWidth { get; set; }
        public static double MySavedProgramHeight { get; set; }



        //Shadows//
        public static string MyShadowColor { get; set; }
		public static int MyShadowDirection { get; set; }
		public static int MyShadowDepth { get; set; }
		public static double MyShadowSoftness { get; set; }

		//Text//
		public static double MyTextFontSize { get; set; }
		public static double MyTextBlockFontSize { get; set; }

		//MenuBorder//
		public static string MyMenuBorderBrushColor { get; set; }

		//Corners//
		public static int MyRoundedCorners { get; set; } 

		//Borders//
		public static int MyBorderWidth { get; set; }

        //FixDifference//

        public static int MyWidthDifferenceFix { get; set; }
        public static int MyHeightDifferenceFix { get; set; }

        //MadeUpWidths//
        public static int MyCompleteSecondGridWidths { get; set; }

        //SaveMargin//
        public static int MySaveByValue { get; set; }


        public App()
		{
			//DefaultProgramWidthOrHeight//
			MyProgramMinWidth = 436;
			MyProgramMinHeight = 600;
			MyFirstGridWidth = 420;
            MySecondGridBarWidth = 9;
            MyProgramMinimizedWidth = MyProgramMinWidth + MySecondGridBarWidth;

			//Shadows//
			MyShadowColor = "#447597";//"#447597"
			MyShadowDirection = -50;
			MyShadowDepth = 5;
			MyShadowSoftness = 0.05;

			//Text//
			MyTextFontSize = 13.5;
			MyTextBlockFontSize = 11.5;

			//MenuBorder//
			MyMenuBorderBrushColor = "#0054A6";//"#0054A6"

			//Corners//
			MyRoundedCorners = 6;

			//Borders//
			MyBorderWidth = 3;

            //FixDifference//
            MyWidthDifferenceFix = 1;
            MyHeightDifferenceFix = 38;

            //MadeUpWidths//
            MyCompleteSecondGridWidths = MySecondGridBarWidth + (MyBorderWidth * 2) + MyWidthDifferenceFix;

            //SaveMargin//
            MySaveByValue = 10;

        }
	}
}
