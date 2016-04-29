using Aspose.Cells;

namespace FluentCells
{
	public static class WorksheetExt
	{
		public static Worksheet Bootstrap(this Worksheet ws)
		{

			ws
				.SetErrorCheck(false)
				.SetStandardHeight()
				.SetStandardWidth();

			ws.PageSetup
				.SetTopMargin(Settings.TopMargin)
				.SetRightMargin(Settings.RightMargin)
				.SetBottomMargin(Settings.BottomMargin)
				.SetLeftMargin(Settings.LeftMargin)
				.SetHeaderMargin(Settings.HeaderMargin)
				.SetFooterMargin(Settings.FooterMargin)
				.SetHFAlignMargins(Settings.IsHFAlignMargins)
				.SetOrientation(Settings.OrientationType);

			return ws;
		}
		public static Worksheet SetErrorCheck(this Worksheet ws, bool isCheck)
		{
			var opts = ws.ErrorCheckOptions;
			var i = opts.Add();
			ErrorCheckOption opt = opts[i];
			opt.SetErrorCheck(ErrorCheckType.EmptyCellRef, isCheck);
			opt.SetErrorCheck(ErrorCheckType.TextDate, isCheck);
			opt.SetErrorCheck(ErrorCheckType.TextNumber, isCheck);
			opt.AddRange(CellArea.CreateCellArea(0, 0, Settings.MaxRows, Settings.MaxColumns));
			return ws;
		}
		public static Worksheet SetName(this Worksheet ws, string sheetName, params object[] args)
		{
			var name = string.Format(sheetName, args);

			foreach (var c in new string[] { "\\", "/", "?", "*", "[", "]", ":", "'" })
				name = name.Replace(c, "_");

			if (name.Length > 30)
				name = name.Substring(0, 30);

			ws.Name = name;
			return ws;
		}
		public static Worksheet SetStandardHeight(this Worksheet ws, double value = Settings.StandardHeight)
		{
			ws.Cells.StandardHeight = value;
			return ws;
		}
		public static Worksheet SetStandardWidth(this Worksheet ws, double value = Settings.StandardWidth)
		{
			ws.Cells.StandardWidth = value;
			return ws;
		}
	}
}
