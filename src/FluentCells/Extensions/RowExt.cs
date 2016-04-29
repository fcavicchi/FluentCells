using Aspose.Cells;
using System.Drawing;

namespace FluentCells
{
	public static class RowExt
	{

		public static Row SetFontBold(this Row row, bool value = true)
		{
			var style = row.Style;
			style.Font.IsBold = value;
			row.ApplyStyle(style, new StyleFlag { FontBold = true });
			return row;
		}

		public static Row SetFontColor(this Row row, Color value)
		{
			var style = row.Style;
			style.Font.Color = value;
			row.ApplyStyle(style, new StyleFlag { FontColor = true });
			return row;
		}

		public static Row SetFontItalic(this Row row, bool value = true)
		{
			var style = row.Style;
			style.Font.IsItalic = value;
			row.ApplyStyle(style, new StyleFlag { FontItalic = true });
			return row;
		}

		public static Row SetFontName(this Row row, string value = Settings.FontName)
		{
			var style = row.Style;
			style.Font.Name = value;
			row.ApplyStyle(style, new StyleFlag { FontName = true });
			return row;
		}

		public static Row SetFontSize(this Row row, int value = Settings.FontSize)
		{
			var style = row.Style;
			style.Font.Size = value;
			row.ApplyStyle(style, new StyleFlag { FontSize = true });
			return row;
		}

		public static Row SetFontStrikeout(this Row row, bool value = true)
		{
			var style = row.Style;
			style.Font.IsStrikeout = value;
			row.ApplyStyle(style, new StyleFlag { FontStrike = true });
			return row;
		}

		public static Row SetFontUnderline(this Row row, FontUnderlineType value)
		{
			var style = row.Style;
			style.Font.Underline = value;
			row.ApplyStyle(style, new StyleFlag { FontUnderline = true });
			return row;
		}

		public static Row SetFontUnderlineToNone(this Row row)
		{
			return row.SetFontUnderline(FontUnderlineType.None);
		}

		public static Row SetFontUnderlineToSingle(this Row row)
		{
			return row.SetFontUnderline(FontUnderlineType.Single);
		}

		public static Row SetForegroundColor(this Row row, System.Drawing.Color value)
		{
			var style = row.Style;
			style.ForegroundColor = value;
			style.Pattern = BackgroundType.Solid;
			row.ApplyStyle(style, new StyleFlag { CellShading = true });
			return row;
		}

		public static Row SetForegroundColorToTransparent(this Row row)
		{
			return row.SetForegroundColor(Color.Transparent);
		}

		public static Row SetFormat(this Row row, DisplayFormat value)
		{
			return row.SetFormat((int)value);
		}

		public static Row SetFormat(this Row row, int value)
		{
			var style = row.Style;
			style.Number = value;
			row.ApplyStyle(style, new StyleFlag { NumberFormat = true });
			return row;
		}

		public static Row SetFormat(this Row row, string value, params object[] args)
		{
			var style = row.Style;
			style.Custom = string.Format(value, args);
			row.ApplyStyle(style, new StyleFlag { NumberFormat = true });
			return row;
		}

		public static Row SetHorizontalAlignment(this Row row, TextAlignmentType value)
		{
			var style = row.Style;
			style.HorizontalAlignment = value;
			row.ApplyStyle(style, new StyleFlag { HorizontalAlignment = true });
			return row;
		}

		public static Row SetHorizontalAlignmentToCenter(this Row row)
		{
			return row.SetHorizontalAlignment(TextAlignmentType.Center);
		}

		public static Row SetHorizontalAlignmentToLeft(this Row row)
		{
			return row.SetHorizontalAlignment(TextAlignmentType.Left);
		}

		public static Row SetHorizontalAlignmentToRight(this Row row)
		{
			return row.SetHorizontalAlignment(TextAlignmentType.Right);
		}

		public static Row SetRowHeight(this Row row, double height)
		{
			row.Height = height;
			return row;
		}

		public static Row SetTextWrapped(this Row row, bool value = true)
		{
			var style = row.Style;
			style.IsTextWrapped = value;
			row.ApplyStyle(style, new StyleFlag { WrapText = true });
			return row;
		}

		public static Row SetVerticalAlignment(this Row row)
		{
			return row.SetVerticalAlignment(TextAlignmentType.Center);
		}

		public static Row SetVerticalAlignment(this Row row, TextAlignmentType value)
		{
			var style = row.Style;
			style.VerticalAlignment = value;
			row.ApplyStyle(style, new StyleFlag { VerticalAlignment = true });
			return row;
		}

		public static Row SetVerticalAlignmentToBottom(this Row row)
		{
			return row.SetVerticalAlignment(TextAlignmentType.Bottom);
		}

		public static Row SetVerticalAlignmentToCenter(this Row row)
		{
			return row.SetVerticalAlignment(TextAlignmentType.Center);
		}

		public static Row SetVerticalAlignmentToTop(this Row row)
		{
			return row.SetVerticalAlignment(TextAlignmentType.Top);
		}

	}
}
