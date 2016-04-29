using Aspose.Cells;
using System.Collections.Generic;
using System.Drawing;

namespace FluentCells
{
	public static class ColumnCollectionExt
	{

		public static IList<Column> SetFontBold(this IList<Column> columns, bool value = true, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetFontBold(value);
			return columns;
		}

		public static IList<Column> SetFontItalic(this IList<Column> columns, bool value = true, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetFontItalic(value);
			return columns;
		}

		public static IList<Column> SetFontName(this IList<Column> columns, string value = Settings.FontName, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetFontName(value);
			return columns;
		}

		public static IList<Column> SetFontSize(this IList<Column> columns, int value = Settings.FontSize, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetFontSize(value);
			return columns;
		}

		public static IList<Column> SetFontUnderline(this IList<Column> columns, FontUnderlineType value, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetFontUnderline(value);
			return columns;
		}

		public static IList<Column> SetFontUnderlineToNone(this IList<Column> columns, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetFontUnderlineToNone();
			return columns;
		}

		public static IList<Column> SetFontUnderlineToSingle(this IList<Column> columns, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetFontUnderlineToSingle();
			return columns;
		}

		public static IList<Column> SetForegroundColor(this IList<Column> columns, Color value, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetForegroundColor(value);
			return columns;
		}

		public static IList<Column> SetForegroundColor(this IList<Column> columns, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetForegroundColorToTransparent();
			return columns;
		}

		public static IList<Column> SetFormat(this IList<Column> columns, DisplayFormat value, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetFormat(value);
			return columns;
		}

		public static IList<Column> SetFormat(this IList<Column> columns, int value, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetFormat(value);
			return columns;
		}

		public static IList<Column> SetHorizontalAlignment(this IList<Column> columns, TextAlignmentType value, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetHorizontalAlignment(value);
			return columns;
		}

		public static IList<Column> SetHorizontalAlignmentToCenter(this IList<Column> columns, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetHorizontalAlignmentToCenter();
			return columns;
		}

		public static IList<Column> SetHorizontalAlignmentToLeft(this IList<Column> columns, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetHorizontalAlignmentToLeft();
			return columns;
		}

		public static IList<Column> SetHorizontalAlignmentToRight(this IList<Column> columns, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetHorizontalAlignmentToRight();
			return columns;
		}

		public static IList<Column> SetTextWrapped(this IList<Column> columns, bool value = true, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetTextWrapped(value);
			return columns;
		}

		public static IList<Column> SetVerticalAlignment(this IList<Column> columns, TextAlignmentType value, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetVerticalAlignment(value);
			return columns;
		}

		public static IList<Column> SetVerticalAlignmentToBottom(this IList<Column> columns, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetVerticalAlignmentToBottom();
			return columns;
		}

		public static IList<Column> SetVerticalAlignmentToCenter(this IList<Column> columns, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetVerticalAlignmentToCenter();
			return columns;
		}

		public static IList<Column> SetVerticalAlignmentToTop(this IList<Column> columns, params int[] columnIndexes)
		{
			foreach (var i in columnIndexes)
				columns[i].SetVerticalAlignmentToTop();
			return columns;
		}

	}
}
