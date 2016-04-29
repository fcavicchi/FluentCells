using Aspose.Cells;
using System.Collections.Generic;

namespace FluentCells
{
	public static class HorizontalPageBreakCollectionExt
	{
		public static IList<HorizontalPageBreak> RemoveLast(this IList<HorizontalPageBreak> breaks)
		{
			var index = breaks.Count - 1;
			if (index >= 0)
				breaks.RemoveAt(index);
			return breaks;
		}
	}
}
