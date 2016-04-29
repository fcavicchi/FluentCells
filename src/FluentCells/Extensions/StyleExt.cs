using Aspose.Cells;
using System.Drawing;

namespace FluentCells
{
	public static class StyleExt
	{

		public static Style SetFontBold(this Style style, bool value = true)
		{
			style.Font.IsBold = value;
			return style;
		}

		public static Style SetFontColor(this Style style, Color value)
		{
			style.Font.Color = value;
			return style;
		}

		public static Style SetFontItalic(this Style style, bool value = true)
		{
			style.Font.IsItalic = value;
			return style;
		}

		public static Style SetFontName(this Style style, string value = Settings.FontName)
		{
			style.Font.Name = value;
			return style;
		}

		public static Style SetFontSize(this Style style, int value = Settings.FontSize)
		{
			style.Font.Size = value;
			return style;
		}

		public static Style SetFontStrikeout(this Style style, bool value = true)
		{
			style.Font.IsStrikeout = value;
			return style;
		}

		public static Style SetFontUnderline(this Style style, FontUnderlineType value)
		{
			style.Font.Underline = value;
			return style;
		}

		public static Style SetFontUnderlineToNone(this Style style)
		{
			return style.SetFontUnderline(FontUnderlineType.None);
		}

		public static Style SetFontUnderlineToSingle(this Style style)
		{
			return style.SetFontUnderline(FontUnderlineType.Single);
		}

		public static Style SetForegroundColor(this Style style, System.Drawing.Color value)
		{
			style.ForegroundColor = value;
			style.Pattern = BackgroundType.Solid;
			return style;
		}

		public static Style SetForegroundColorToTransparent(this Style style)
		{
			return style.SetForegroundColor(Color.Transparent);
		}

		public static Style SetFormat(this Style style, DisplayFormat value)
		{
			return style.SetFormat((int)value);
		}

		public static Style SetFormat(this Style style, int value)
		{
			style.Number = value;
			return style;
		}

		public static Style SetFormat(this Style style, string value, params object[] args)
		{
			style.Custom = string.Format(value, args);
			return style;
		}

		public static Style SetHorizontalAlignment(this Style style, TextAlignmentType value)
		{
			style.HorizontalAlignment = value;
			return style;
		}

		public static Style SetHorizontalAlignmentToCenter(this Style style)
		{
			return style.SetHorizontalAlignment(TextAlignmentType.Center);
		}

		public static Style SetHorizontalAlignmentToLeft(this Style style)
		{
			return style.SetHorizontalAlignment(TextAlignmentType.Left);
		}

		public static Style SetHorizontalAlignmentToRight(this Style style)
		{
			return style.SetHorizontalAlignment(TextAlignmentType.Right);
		}

		public static Style SetTextWrapped(this Style style, bool value = true)
		{
			style.IsTextWrapped = value;
			return style;
		}

		public static Style SetVerticalAlignment(this Style style, TextAlignmentType value)
		{
			style.VerticalAlignment = value;
			return style;
		}

		public static Style SetVerticalAlignmentToBottom(this Style style)
		{
			return style.SetVerticalAlignment(TextAlignmentType.Bottom);
		}

		public static Style SetVerticalAlignmentToCenter(this Style style)
		{
			return style.SetVerticalAlignment(TextAlignmentType.Center);
		}

		public static Style SetVerticalAlignmentToTop(this Style style)
		{
			return style.SetVerticalAlignment(TextAlignmentType.Top);
		}

	}
}
