using Aspose.Cells;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace FluentCells.Tests
{

	[TestClass]
	public class CellTests
	{
		private static int c = 0;
		private static int r = 0;
		private static Workbook wb;
		private static Worksheet ws;
		private Cell cell;

		[ClassInitialize]
		public static void ClassInitialize(TestContext context)
		{
			wb = new Workbook();
			wb.Bootstrap();
			ws = wb.Worksheets.Add("CellTests");
			ws.Bootstrap();
		}

		[TestInitialize]
		public void TestInitialize()
		{
			cell = ws.Cells[r, c++];
			cell.Put(TestContext.TestName);
		}

		[TestCleanup]
		public void TestCleanup()
		{
			r++;
			c = 0;
		}

		[ClassCleanup]
		public static void ClassCleanup()
		{
			ws.AutoFitColumns();
			var fileName = @"C:\CellTests.xlsx";
			wb.Save(fileName);
		}

		[TestMethod]
		public void Test_Cell_SetFontBold()
		{
			var values = new List<bool> { true, false };

			foreach (var value in values) {
				cell = ws.Cells[r, c++];
				cell.PutValue("Font bold");
				cell.SetFontBold(value);
				Assert.IsTrue(cell.GetStyle().Font.IsBold == value);
			}
		}

		[TestMethod]
		public void Test_Cell_SetFontColor()
		{
			var colors = new Color[] { Color.Blue, Color.Green, Color.Red, Color.Yellow };

			foreach (var color in colors) {
				cell = ws.Cells[r, c++];
				cell.Put("Font color {0}", color.Name);
				cell.SetFontColor(color);
				var newColor = cell.GetStyle().Font.Color;
				Assert.IsTrue(newColor.R == color.R);
				Assert.IsTrue(newColor.G == color.G);
				Assert.IsTrue(newColor.B == color.B);
			}
		}

		[TestMethod]
		public void Test_Cell_SetFontItalic()
		{
			var values = new List<bool> { true, false };

			foreach (var value in values) {
				cell = ws.Cells[r, c++];
				cell.Put("Font italic");
				cell.SetFontItalic(value);
				Assert.IsTrue(cell.GetStyle().Font.IsItalic == value);
			}
		}

		[TestMethod]
		public void Test_Cell_SetFontName()
		{
			var fonts = new string[] { "Arial", "Calibri", "Consolas", "Tahoma", "Verdana" };

			foreach (var font in fonts) {
				cell = ws.Cells[r, c++];
				cell.Put("Font name {0}", font);
				cell.SetFontName(font);
				Assert.IsTrue(cell.GetStyle().Font.Name == font);
			}
		}

		[TestMethod]
		public void Test_Cell_SetFontSize()
		{
			var sizes = new int[] { 8, 10, 12, 14, 16 };

			foreach (var size in sizes) {
				cell = ws.Cells[r, c++];
				cell.Put("Font size {0}", size);
				cell.SetFontSize(size);
				Assert.IsTrue(cell.GetStyle().Font.Size == size);
			}
		}

		[TestMethod]
		public void Test_Cell_SetFontStrikeout()
		{
			var values = new List<bool> { true, false };

			foreach (var value in values) {
				cell = ws.Cells[r, c++];
				cell.PutValue("Font strikeout");
				cell.SetFontStrikeout(value);
				Assert.IsTrue(cell.GetStyle().Font.IsStrikeout == value);
			}
		}

		[TestMethod]
		public void Test_Cell_SetFontUnderline()
		{
			var values = new List<FontUnderlineType> { FontUnderlineType.Accounting, FontUnderlineType.Double, FontUnderlineType.DoubleAccounting, FontUnderlineType.None, FontUnderlineType.Single };

			foreach (var value in values) {
				cell = ws.Cells[r, c++];
				cell.PutValue(Enum.GetName(typeof(FontUnderlineType), value));
				cell.SetFontUnderline(value);
				Assert.IsTrue(cell.GetStyle().Font.Underline == value);
			}
		}

		[TestMethod]
		public void Test_Cell_SetForegroudColor()
		{
			var colors = new Color[] { Color.Blue, Color.Green, Color.Red, Color.Yellow };

			foreach (var color in colors) {
				cell = ws.Cells[r, c++];
				cell.Put("Foregroud color {0}", color.Name);
				cell.SetForegroundColor(color);
				var newColor = cell.GetStyle().ForegroundColor;
				Assert.IsTrue(newColor.R == color.R);
				Assert.IsTrue(newColor.G == color.G);
				Assert.IsTrue(newColor.B == color.B);
			}
		}

		[TestMethod]
		public void Test_Cell_SetFormat()
		{
			foreach (DisplayFormat value in Enum.GetValues(typeof(DisplayFormat))) {
				cell = ws.Cells[r, c++];
				cell.PutValue(Enum.GetName(typeof(DisplayFormat), value));
				cell.SetFormat(value);
				Assert.IsTrue(cell.GetStyle().Number == (int)value);
			}
		}

		[TestMethod]
		public void Test_Cell_SetHorizontalAlignment()
		{
			var values = new List<TextAlignmentType> { TextAlignmentType.Center, TextAlignmentType.Left, TextAlignmentType.Right };

			foreach (var value in values) {
				cell = ws.Cells[r, c++];
				cell.PutValue(Enum.GetName(typeof(TextAlignmentType), value));
				cell.SetHorizontalAlignment(value);
				Assert.IsTrue(cell.GetStyle().HorizontalAlignment == value);
			}
		}

		[TestMethod]
		public void Test_Cell_SetTextWrapped()
		{
			var values = new List<bool> { true, false };

			foreach (var value in values) {
				cell = ws.Cells[r, c++];
				cell.PutValue("Text wrapped");
				cell.SetTextWrapped(value);
				Assert.IsTrue(cell.GetStyle().IsTextWrapped == value);
			}
		}

		[TestMethod]
		public void Test_Cell_SetVerticalAlignment()
		{
			var values = new List<TextAlignmentType> { TextAlignmentType.Bottom, TextAlignmentType.Center, TextAlignmentType.Top };

			foreach (var value in values) {
				cell = ws.Cells[r, c++];
				cell.PutValue(Enum.GetName(typeof(TextAlignmentType), value));
				cell.SetVerticalAlignment(value);
				Assert.IsTrue(cell.GetStyle().VerticalAlignment == value);
			}
		}

		public TestContext TestContext { get; set; }

	}
}
