using Aspose.Cells;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FluentCells.Tests {

	[TestClass]
	public class WorkbookTests {
		private static Workbook wb;

		[ClassInitialize]
		public static void ClassInitialize(TestContext context) {
			wb = new Workbook();
		}

		[TestInitialize]
		public void TestInitialize() { }

		[TestCleanup]
		public void TestCleanup() { }

		[ClassCleanup]
		public static void ClassCleanup() { }

		[TestMethod]
		public void Test_Workbook_Bootstrap() {
			wb.Bootstrap();
			Assert.IsTrue(wb.Worksheets.Count == 0);
			var style = wb.DefaultStyle;
			Assert.AreEqual(style.Font.Name, Settings.FontName);
			Assert.AreEqual(style.Font.Size, Settings.FontSize);
			Assert.AreEqual(style.HorizontalAlignment, TextAlignmentType.Left);
			Assert.AreEqual(style.VerticalAlignment, TextAlignmentType.Center);
		}

		[TestMethod]
		public void Test_Workbook_SetFileFormat() {
			var oldFileFormat = wb.FileFormat;
			var newFileFormat = FileFormatType.Xltx;
			wb.SetFileFormat(newFileFormat);
			Assert.AreNotEqual(wb.FileFormat, oldFileFormat);
		}

		[TestMethod]
		public void Test_Workbook_SetFileName() {
			var oldFileName = wb.FileName;
			var newFileName = string.Format("{0}_NEW", oldFileName);
			wb.SetFileName(newFileName);
			Assert.AreNotEqual(wb.FileName, oldFileName);
		}

		public TestContext TestContext { get; set; }

	}
}
