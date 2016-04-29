using Aspose.Cells;
using System.Collections.Generic;

namespace FluentCells
{
	public static class WorkbookExt
	{

		public static Workbook Bootstrap(this Workbook wb)
		{
			wb.Worksheets
				.Clear();

			var style = wb
				.DefaultStyle
				.SetFontName(Settings.FontName)
				.SetFontSize(Settings.FontSize)
				.SetHorizontalAlignmentToLeft()
				.SetVerticalAlignmentToCenter();

			wb.DefaultStyle = style;
			return wb;
		}

		public static IList<Style> GetStyles(this Workbook wb)
		{
			IList<Style> styles = new List<Style>();
			for (var i = 0; i < wb.CountOfStylesInPool; i++)
				styles.Add(wb.GetStyleInPool(i));
			return styles;
		}

		public static Workbook SetFileFormat(this Workbook wb, FileFormatType fileFormat)
		{
			wb.FileFormat = fileFormat;
			return wb;
		}

		public static Workbook SetFileName(this Workbook wb, string fileName)
		{
			wb.FileName = fileName;
			return wb;
		}

	}
}
