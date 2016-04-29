using Aspose.Cells;
using System.Drawing;

namespace FluentCells
{
	public static class RangeExt
	{

		public static Range SetBorder(this Range range, BorderType borderType)
		{
			range.SetOutlineBorder(borderType, CellBorderType.Thin, Color.Silver);
			return range;
		}

		public static Range SetColumnWidth(this Range range, double width)
		{
			range.ColumnWidth = width;
			return range;
		}

		public static Range SetFontBold(this Range range, bool value = true)
		{
			var style = new Style();
			style.Font.IsBold = value;
			range.ApplyStyle(style, new StyleFlag { FontBold = true });
			return range;
		}

		public static Range SetFontColor(this Range range, Color value)
		{
			var style = new Style();
			style.Font.Color = value;
			range.ApplyStyle(style, new StyleFlag { FontColor = true });
			return range;
		}

		public static Range SetFontItalic(this Range range, bool value = true)
		{
			var style = new Style();
			style.Font.IsItalic = value;
			range.ApplyStyle(style, new StyleFlag { FontItalic = true });
			return range;
		}

		public static Range SetFontName(this Range range, string value = Settings.FontName)
		{
			var style = new Style();
			style.Font.Name = value;
			range.ApplyStyle(style, new StyleFlag { FontName = true });
			return range;
		}

		public static Range SetFontSize(this Range range, int value = Settings.FontSize)
		{
			var style = new Style();
			style.Font.Size = value;
			range.ApplyStyle(style, new StyleFlag { FontSize = true });
			return range;
		}

		public static Range SetFontStrikeout(this Range range, bool value = true)
		{
			var style = new Style();
			style.Font.IsStrikeout = value;
			range.ApplyStyle(style, new StyleFlag { FontStrike = true });
			return range;
		}

		public static Range SetFontUnderline(this Range range, FontUnderlineType value)
		{
			var style = new Style();
			style.Font.Underline = value;
			range.ApplyStyle(style, new StyleFlag { FontUnderline = true });
			return range;
		}

		public static Range SetFontUnderlineToNone(this Range range)
		{
			return range.SetFontUnderline(FontUnderlineType.None);
		}

		public static Range SetFontUnderlineToSingle(this Range range)
		{
			return range.SetFontUnderline(FontUnderlineType.Single);
		}

		public static Range SetForegroundColor(this Range range, System.Drawing.Color value)
		{
			var style = new Style();
			style.ForegroundColor = value;
			style.Pattern = BackgroundType.Solid;
			range.ApplyStyle(style, new StyleFlag { CellShading = true });
			return range;
		}

		public static Range SetForegroundColorToTransparent(this Range range)
		{
			return range.SetForegroundColor(Color.Transparent);
		}

		public static Range SetFormat(this Range range, DisplayFormat value)
		{
			return range.SetFormat((int)value);
		}

		public static Range SetFormat(this Range range, int value)
		{
			var style = new Style();
			style.Number = value;
			range.ApplyStyle(style, new StyleFlag { NumberFormat = true });
			return range;
		}

		public static Range SetFormat(this Range range, string value, params object[] args)
		{
			var style = new Style();
			style.Custom = string.Format(value, args);
			range.ApplyStyle(style, new StyleFlag { NumberFormat = true });
			return range;
		}

		public static Range SetHorizontalAlignment(this Range range, TextAlignmentType value)
		{
			var style = new Style();
			style.HorizontalAlignment = value;
			range.ApplyStyle(style, new StyleFlag { HorizontalAlignment = true });
			return range;
		}

		public static Range SetHorizontalAlignmentToCenter(this Range range)
		{
			return range.SetHorizontalAlignment(TextAlignmentType.Center);
		}

		public static Range SetHorizontalAlignmentToLeft(this Range range)
		{
			return range.SetHorizontalAlignment(TextAlignmentType.Left);
		}

		public static Range SetHorizontalAlignmentToRight(this Range range)
		{
			return range.SetHorizontalAlignment(TextAlignmentType.Right);
		}

		public static Range SetRowHeight(this Range range, double height)
		{
			range.RowHeight = height;
			return range;
		}

		public static Range SetTextWrapped(this Range range, bool value = true)
		{
			var style = new Style();
			style.IsTextWrapped = value;
			range.ApplyStyle(style, new StyleFlag { WrapText = true });
			return range;
		}

		public static Range SetVerticalAlignment(this Range range, TextAlignmentType value)
		{
			var style = new Style();
			style.VerticalAlignment = value;
			range.ApplyStyle(style, new StyleFlag { VerticalAlignment = true });
			return range;
		}

		public static Range SetVerticalAlignmentToBottom(this Range range)
		{
			return range.SetVerticalAlignment(TextAlignmentType.Bottom);
		}

		public static Range SetVerticalAlignmentToCenter(this Range range)
		{
			return range.SetVerticalAlignment(TextAlignmentType.Center);
		}

		public static Range SetVerticalAlignmentToTop(this Range range)
		{
			return range.SetVerticalAlignment(TextAlignmentType.Top);
		}

	}
}
