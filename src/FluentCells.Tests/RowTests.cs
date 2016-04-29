using Aspose.Cells;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Drawing;

namespace FluentCells.Tests {

	[TestClass]
	public class RowTests {
		private static int c = 0;
		private static int r = 0;
		private static Workbook wb;
		private static Worksheet ws;
		private Row row;

		[ClassInitialize]
		public static void ClassInitialize(TestContext context) {
			wb = new Workbook();
			wb.Bootstrap();
			ws = wb.Worksheets.Add("RowTests");
			ws.Bootstrap();
		}

		[TestInitialize]
		public void TestInitialize() {
			//
		}

		[TestCleanup]
		public void TestCleanup() {
			//
		}

		[ClassCleanup]
		public static void ClassCleanup() {
			ws.AutoFitColumns();
			var fileName = @"C:\RowTests.xlsx";
			wb.Save(fileName);
		}

		[TestMethod]
		public void Test_Row_SetFontBold() {
			var values = new List<bool> { true, false };

			foreach (var value in values) {
				c = 0;
				row = ws.Cells[r, c].GetRow();
				row.SetFontBold(value);
				row[c++].Put(TestContext.TestName);
				row[c++].Put(value);
				Assert.IsTrue(row.Style.Font.IsBold == value);
				r++;
			}
		}

		[TestMethod]
		public void Test_Row_SetFontColor() {
			var colors = new Color[] { Color.Blue, Color.Green, Color.Red, Color.Yellow };

			foreach (var color in colors) {
				c = 0;
				row = ws.Cells[r, c].GetRow();
				row.SetFontColor(color);
				row[c++].Put(TestContext.TestName);
				row[c++].Put(color.Name);
				var newColor = row.Style.Font.Color;
				Assert.IsTrue(newColor.R == color.R);
				Assert.IsTrue(newColor.G == color.G);
				Assert.IsTrue(newColor.B == color.B);
				r++;
			}
		}

		[TestMethod]
		public void Test_Row_SetFontItalic() {
			var values = new List<bool> { true, false };

			foreach (var value in values) {
				c = 0;
				row = ws.Cells[r, c].GetRow();
				row.SetFontItalic(value);
				row[c++].Put(TestContext.TestName);
				row[c++].Put(value);
				Assert.IsTrue(row.Style.Font.IsItalic == value);
				r++;
			}
		}

		public TestContext TestContext { get; set; }

	}
}
