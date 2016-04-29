using Aspose.Cells;
using System.Drawing;

namespace FluentCells
{
	public static class ColumnExt
	{

		public static Column SetColumnWidth(this Column column, double width)
		{
			column.Width = width;
			return column;
		}

		public static Column SetFontBold(this Column column, bool value = true)
		{
			var style = column.Style;
			style.Font.IsBold = value;
			column.ApplyStyle(style, new StyleFlag { FontBold = true });
			return column;
		}

		public static Column SetFontColor(this Column column, Color value)
		{
			var style = column.Style;
			style.Font.Color = value;
			column.ApplyStyle(style, new StyleFlag { FontColor = true });
			return column;
		}

		public static Column SetFontItalic(this Column column, bool value = true)
		{
			var style = column.Style;
			style.Font.IsItalic = value;
			column.ApplyStyle(style, new StyleFlag { FontItalic = true });
			return column;
		}

		public static Column SetFontName(this Column column, string value = Settings.FontName)
		{
			var style = column.Style;
			style.Font.Name = value;
			column.ApplyStyle(style, new StyleFlag { FontName = true });
			return column;
		}

		public static Column SetFontSize(this Column column, int value = Settings.FontSize)
		{
			var style = column.Style;
			style.Font.Size = value;
			column.ApplyStyle(style, new StyleFlag { FontSize = true });
			return column;
		}

		public static Column SetFontStrikeout(this Column column, bool value = true)
		{
			var style = new Style();
			style.Font.IsStrikeout = value;
			column.ApplyStyle(style, new StyleFlag { FontStrike = true });
			return column;
		}

		public static Column SetFontUnderline(this Column column, FontUnderlineType value)
		{
			var style = column.Style;
			style.Font.Underline = value;
			column.ApplyStyle(style, new StyleFlag { FontUnderline = true });
			return column;
		}

		public static Column SetFontUnderlineToNone(this Column column)
		{
			return column.SetFontUnderline(FontUnderlineType.None);
		}

		public static Column SetFontUnderlineToSingle(this Column column)
		{
			return column.SetFontUnderline(FontUnderlineType.Single);
		}

		public static Column SetForegroundColor(this Column column, Color value)
		{
			var style = column.Style;
			style.ForegroundColor = value;
			style.Pattern = BackgroundType.Solid;
			column.ApplyStyle(style, new StyleFlag { CellShading = true });
			return column;
		}

		public static Column SetForegroundColorToTransparent(this Column column)
		{
			return column.SetForegroundColor(Color.Transparent);
		}

		public static Column SetFormat(this Column column, DisplayFormat value)
		{
			return column.SetFormat((int)value);
		}

		public static Column SetFormat(this Column column, int value)
		{
			var style = column.Style;
			style.Number = value;
			column.ApplyStyle(style, new StyleFlag { NumberFormat = true });
			return column;
		}

		public static Column SetFormat(this Column column, string value, params object[] args)
		{
			var style = column.Style;
			style.Custom = string.Format(value, args);
			column.ApplyStyle(style, new StyleFlag { NumberFormat = true });
			return column;
		}

		public static Column SetHorizontalAlignment(this Column column, TextAlignmentType value)
		{
			var style = column.Style;
			style.HorizontalAlignment = value;
			column.ApplyStyle(style, new StyleFlag { HorizontalAlignment = true });
			return column;
		}

		public static Column SetHorizontalAlignmentToCenter(this Column column)
		{
			return column.SetHorizontalAlignment(TextAlignmentType.Center);
		}

		public static Column SetHorizontalAlignmentToLeft(this Column column)
		{
			return column.SetHorizontalAlignment(TextAlignmentType.Left);
		}

		public static Column SetHorizontalAlignmentToRight(this Column column)
		{
			return column.SetHorizontalAlignment(TextAlignmentType.Right);
		}

		public static Column SetTextWrapped(this Column column, bool value = true)
		{
			var style = column.Style;
			style.IsTextWrapped = value;
			column.ApplyStyle(style, new StyleFlag { WrapText = true });
			return column;
		}

		public static Column SetVerticalAlignment(this Column column, TextAlignmentType value)
		{
			var style = column.Style;
			style.VerticalAlignment = value;
			column.ApplyStyle(style, new StyleFlag { VerticalAlignment = true });
			return column;
		}

		public static Column SetVerticalAlignmentToBottom(this Column column)
		{
			return column.SetVerticalAlignment(TextAlignmentType.Bottom);
		}

		public static Column SetVerticalAlignmentToCenter(this Column column)
		{
			return column.SetVerticalAlignment(TextAlignmentType.Center);
		}

		public static Column SetVerticalAlignmentToTop(this Column column)
		{
			return column.SetVerticalAlignment(TextAlignmentType.Top);
		}

	}
}
