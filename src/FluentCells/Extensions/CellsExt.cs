using Aspose.Cells;
using System.Drawing;

namespace FluentCells
{
	public static class CellsExt
	{

		public static Cells SetFontBold(this Cells cells, bool value = true)
		{
			var style = new Style();
			style.Font.IsBold = value;
			cells.ApplyStyle(style, new StyleFlag { FontBold = true });
			return cells;
		}

		public static Cells SetFontColor(this Cells cells, Color value)
		{
			var style = new Style();
			style.Font.Color = value;
			cells.ApplyStyle(style, new StyleFlag { FontColor = true });
			return cells;
		}

		public static Cells SetFontItalic(this Cells cells, bool value = true)
		{
			var style = new Style();
			style.Font.IsItalic = value;
			cells.ApplyStyle(style, new StyleFlag { FontItalic = true });
			return cells;
		}

		public static Cells SetFontName(this Cells cells, string value = Settings.FontName)
		{
			var style = new Style();
			style.Font.Name = value;
			cells.ApplyStyle(style, new StyleFlag { FontName = true });
			return cells;
		}

		public static Cells SetFontSize(this Cells cells, int value = Settings.FontSize)
		{
			var style = new Style();
			style.Font.Size = value;
			cells.ApplyStyle(style, new StyleFlag { FontSize = true });
			return cells;
		}

		public static Cells SetFontStrikeout(this Cells cells, bool value = true)
		{
			var style = new Style();
			style.Font.IsStrikeout = value;
			cells.ApplyStyle(style, new StyleFlag { FontStrike = true });
			return cells;
		}

		public static Cells SetFontUnderline(this Cells cells, FontUnderlineType value)
		{
			var style = new Style();
			style.Font.Underline = value;
			cells.ApplyStyle(style, new StyleFlag { FontUnderline = true });
			return cells;
		}

		public static Cells SetFontUnderlineToNone(this Cells cells)
		{
			return cells.SetFontUnderline(FontUnderlineType.None);
		}

		public static Cells SetFontUnderlineToSingle(this Cells cells)
		{
			return cells.SetFontUnderline(FontUnderlineType.Single);
		}

		public static Cells SetForegroundColor(this Cells cells, Color value)
		{
			var style = new Style();
			style.ForegroundColor = value;
			style.Pattern = BackgroundType.Solid;
			cells.ApplyStyle(style, new StyleFlag { CellShading = true });
			return cells;
		}

		public static Cells SetForegroundColorToTransparent(this Cells cells)
		{
			return cells.SetForegroundColor(Color.Transparent);
		}

		public static Cells SetFormat(this Cells cells, DisplayFormat value)
		{
			return cells.SetFormat((int)value);
		}

		public static Cells SetFormat(this Cells cells, int value)
		{
			var style = new Style();
			style.Number = value;
			cells.ApplyStyle(style, new StyleFlag { NumberFormat = true });
			return cells;
		}

		public static Cells SetFormat(this Cells cells, string value, params object[] args)
		{
			var style = new Style();
			style.Custom = string.Format(value, args);
			cells.ApplyStyle(style, new StyleFlag { NumberFormat = true });
			return cells;
		}

		public static Cells SetHorizontalAlignment(this Cells cells, TextAlignmentType value)
		{
			var style = new Style();
			style.HorizontalAlignment = value;
			cells.ApplyStyle(style, new StyleFlag { HorizontalAlignment = true });
			return cells;
		}

		public static Cells SetHorizontalAlignmentToCenter(this Cells cells)
		{
			return cells.SetHorizontalAlignment(TextAlignmentType.Center);
		}

		public static Cells SetHorizontalAlignmentToLeft(this Cells cells)
		{
			return cells.SetHorizontalAlignment(TextAlignmentType.Left);
		}

		public static Cells SetHorizontalAlignmentToRight(this Cells cells)
		{
			return cells.SetHorizontalAlignment(TextAlignmentType.Right);
		}

		public static Cells SetTextWrapped(this Cells cells, bool value = true)
		{
			var style = new Style();
			style.IsTextWrapped = value;
			cells.ApplyStyle(style, new StyleFlag { WrapText = true });
			return cells;
		}

		public static Cells SetVerticalAlignment(this Cells cells, TextAlignmentType value)
		{
			var style = new Style();
			style.VerticalAlignment = value;
			cells.ApplyStyle(style, new StyleFlag { VerticalAlignment = true });
			return cells;
		}

		public static Cells SetVerticalAlignmentToBottom(this Cells cells)
		{
			return cells.SetVerticalAlignment(TextAlignmentType.Bottom);
		}

		public static Cells SetVerticalAlignmentToCenter(this Cells cells)
		{
			return cells.SetVerticalAlignment(TextAlignmentType.Center);
		}

		public static Cells SetVerticalAlignmentToTop(this Cells cells)
		{
			return cells.SetVerticalAlignment(TextAlignmentType.Top);
		}

	}
}
