using Aspose.Cells;
using System.Collections.Generic;

namespace FluentCells
{
	public static class WorksheetCollectionExt
	{
		public static Worksheet Add(this IList<Worksheet> worksheets, string sheetName, params object[] args)
		{
			return worksheets.Add(string.Format(sheetName, args));
		}
	}
}
