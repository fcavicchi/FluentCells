using Aspose.Cells;
using System;
using System.Drawing;

namespace FluentCells
{
	public static class CellExt
	{

		public static Cell AutoFitColumn(this Cell cell)
		{
			cell.Worksheet.AutoFitColumn(cell.Column);
			return cell;
		}

		public static Cell AutoFitRow(this Cell cell)
		{
			cell.Worksheet.AutoFitRow(cell.Row);
			return cell;
		}

		public static Range CreateRange(this Cell cell, int totalRows, int totalColumns)
		{
			return cell.Worksheet.Cells.CreateRange(cell.Row, cell.Column, totalRows, totalColumns);
		}

		public static Range CreateRangeColumn(this Cell cell)
		{
			return cell.CreateRange(Settings.MaxRows, 1);
		}

		public static Range CreateRangeRow(this Cell cell)
		{
			return cell.CreateRange(1, Settings.MaxColumns);
		}

		public static Column GetColumn(this Cell cell)
		{
			return cell.Worksheet.Cells.Columns[cell.Column];
		}

		public static Row GetRow(this Cell cell)
		{
			return cell.Worksheet.Cells.Rows[cell.Row];
		}

		public static Cell Merge(this Cell cell, int totalRows, int totalColumns)
		{
			cell.Worksheet.Cells.Merge(cell.Row, cell.Column, totalRows, totalColumns);
			return cell;
		}

		public static Cell Put(this Cell cell, bool value)
		{
			cell.PutValue(value);
			return cell;
		}

		public static Cell Put(this Cell cell, DateTime value)
		{
			cell.PutValue(value);
			return cell;
		}

		public static Cell Put(this Cell cell, double value)
		{
			cell.PutValue(value);
			return cell;
		}

		public static Cell Put(this Cell cell, decimal value)
		{
			cell.PutValue(value);
			return cell;
		}

		public static Cell Put(this Cell cell, int value)
		{
			cell.PutValue(value);
			return cell;
		}

		public static Cell Put(this Cell cell, long value)
		{
			cell.PutValue(value);
			return cell;
		}

		public static Cell Put(this Cell cell, object value)
		{
			cell.PutValue(value);
			return cell;
		}

		public static Cell Put(this Cell cell, string value)
		{
			cell.PutValue(value);
			return cell;
		}

		public static Cell Put(this Cell cell, string value, params object[] args)
		{
			cell.PutValue(string.Format(value, args));
			return cell;
		}

		public static Cell PutHyperlink(this Cell cell, string address, string textToDisplay, string screenTip = null)
		{
			var w = cell.Worksheet;
			var i = w.Hyperlinks.Add(cell.Row, cell.Column, 1, 1, address);
			var h = w.Hyperlinks[i];
			h.TextToDisplay = textToDisplay;
			h.ScreenTip = screenTip;
			cell.SetFontColor(System.Drawing.Color.DarkBlue);
			cell.SetFontName();
			cell.SetFontSize();
			cell.SetFontUnderlineToNone();
			return cell;
		}

		public static Cell SetColumnWidth(this Cell cell, double width)
		{
			cell.Worksheet.Cells.SetColumnWidth(cell.Row, width);
			return cell;
		}

		public static Cell SetFontBold(this Cell cell, bool value = true)
		{
			var style = cell.GetStyle();
			style.Font.IsBold = value;
			cell.SetStyle(style);
			return cell;
		}

		public static Cell SetFontColor(this Cell cell, Color value)
		{
			var style = cell.GetStyle();
			style.Font.Color = value;
			cell.SetStyle(style);
			return cell;
		}

		public static Cell SetFontItalic(this Cell cell, bool value = true)
		{
			var style = cell.GetStyle();
			style.Font.IsItalic = value;
			cell.SetStyle(style);
			return cell;
		}

		public static Cell SetFontName(this Cell cell, string value = Settings.FontName)
		{
			var style = cell.GetStyle();
			style.Font.Name = value;
			cell.SetStyle(style);
			return cell;
		}

		public static Cell SetFontSize(this Cell cell, int value = Settings.FontSize)
		{
			var style = cell.GetStyle();
			style.Font.Size = value;
			cell.SetStyle(style);
			return cell;
		}

		public static Cell SetFontStrikeout(this Cell cell, bool value = true)
		{
			var style = cell.GetStyle();
			style.Font.IsStrikeout = value;
			cell.SetStyle(style);
			return cell;
		}

		public static Cell SetFontUnderline(this Cell cell, FontUnderlineType value)
		{
			var style = cell.GetStyle();
			style.Font.Underline = value;
			cell.SetStyle(style);
			return cell;
		}

		public static Cell SetFontUnderlineToNone(this Cell cell)
		{
			return cell.SetFontUnderline(FontUnderlineType.None);
		}

		public static Cell SetFontUnderlineToSingle(this Cell cell)
		{
			return cell.SetFontUnderline(FontUnderlineType.Single);
		}

		public static Cell SetForegroundColor(this Cell cell, Color value)
		{
			var style = cell.GetStyle();
			style.ForegroundColor = value;
			style.Pattern = BackgroundType.Solid;
			cell.SetStyle(style);
			return cell;
		}

		public static Cell SetForegroundColorToTransparent(this Cell cell)
		{
			return cell.SetForegroundColor(Color.Transparent);
		}

		public static Cell SetFormat(this Cell cell, DisplayFormat value)
		{
			return cell.SetFormat((int)value);
		}

		public static Cell SetFormat(this Cell cell, int value)
		{
			var style = cell.GetStyle();
			style.Number = value;
			cell.SetStyle(style);
			return cell;
		}

		public static Cell SetFormat(this Cell cell, string value, params object[] args)
		{
			var style = cell.GetStyle();
			style.Custom = string.Format(value, args);
			cell.SetStyle(style);
			return cell;
		}

		public static Cell SetHorizontalAlignment(this Cell cell, TextAlignmentType value)
		{
			var style = cell.GetStyle();
			style.HorizontalAlignment = value;
			cell.SetStyle(style);
			return cell;
		}

		public static Cell SetHorizontalAlignmentToCenter(this Cell cell)
		{
			return cell.SetHorizontalAlignment(TextAlignmentType.Center);
		}

		public static Cell SetHorizontalAlignmentToLeft(this Cell cell)
		{
			return cell.SetHorizontalAlignment(TextAlignmentType.Left);
		}

		public static Cell SetHorizontalAlignmentToRight(this Cell cell)
		{
			return cell.SetHorizontalAlignment(TextAlignmentType.Right);
		}

		public static Cell SetRowHeight(this Cell cell, double height)
		{
			cell.Worksheet.Cells.SetRowHeight(cell.Row, height);
			return cell;
		}

		public static Cell SetTextWrapped(this Cell cell, bool value = true)
		{
			var style = cell.GetStyle();
			style.IsTextWrapped = value;
			cell.SetStyle(style);
			return cell;
		}

		public static Cell SetVerticalAlignment(this Cell cell, TextAlignmentType value)
		{
			var style = cell.GetStyle();
			style.VerticalAlignment = value;
			cell.SetStyle(style);
			return cell;
		}

		public static Cell SetVerticalAlignmentToBottom(this Cell cell)
		{
			return cell.SetVerticalAlignment(TextAlignmentType.Bottom);
		}

		public static Cell SetVerticalAlignmentToCenter(this Cell cell)
		{
			return cell.SetVerticalAlignment(TextAlignmentType.Center);
		}

		public static Cell SetVerticalAlignmentToTop(this Cell cell)
		{
			return cell.SetVerticalAlignment(TextAlignmentType.Top);
		}

	}
}
