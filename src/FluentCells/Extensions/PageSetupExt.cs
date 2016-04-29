using Aspose.Cells;

namespace FluentCells
{
	public static class PageSetupExt
	{
		public static PageSetup SetBottomMargin(this PageSetup pageSetup, double value)
		{
			pageSetup.BottomMargin = value;
			return pageSetup;
		}
		public static PageSetup SetFooter(this PageSetup pageSetup, int section, string footerScript, params object[] args)
		{
			pageSetup.SetFooter(section, string.Format(footerScript, args));
			return pageSetup;
		}
		public static PageSetup SetFooterMargin(this PageSetup pageSetup, double value)
		{
			pageSetup.FooterMargin = value;
			return pageSetup;
		}
		public static PageSetup SetHeader(this PageSetup pageSetup, int section, string headerScript, params object[] args)
		{
			pageSetup.SetHeader(section, string.Format(headerScript, args));
			return pageSetup;
		}
		public static PageSetup SetHeaderMargin(this PageSetup pageSetup, double value)
		{
			pageSetup.HeaderMargin = value;
			return pageSetup;
		}
		public static PageSetup SetHFAlignMargins(this PageSetup pageSetup, bool value)
		{
			pageSetup.IsHFAlignMargins = value;
			return pageSetup;
		}
		public static PageSetup SetLeftMargin(this PageSetup pageSetup, double value)
		{
			pageSetup.LeftMargin = value;
			return pageSetup;
		}
		public static PageSetup SetOrientation(this PageSetup pageSetup, PageOrientationType value)
		{
			pageSetup.Orientation = value;
			return pageSetup;
		}
		public static PageSetup SetOrientationToLandscape(this PageSetup pageSetup)
		{
			return pageSetup.SetOrientation(PageOrientationType.Landscape);
		}
		public static PageSetup SetOrientationToPortrait(this PageSetup pageSetup)
		{
			return pageSetup.SetOrientation(PageOrientationType.Portrait);
		}
		public static PageSetup SetRightMargin(this PageSetup pageSetup, double value)
		{
			pageSetup.RightMargin = value;
			return pageSetup;
		}
		public static PageSetup SetTopMargin(this PageSetup pageSetup, double value)
		{
			pageSetup.TopMargin = value;
			return pageSetup;
		}
	}
}
