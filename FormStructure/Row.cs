using Microsoft.Office.Interop.Excel;
using System;
using System.Xml.Serialization;

namespace БАРСШаблон
{
	[Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class Row
	{
		public Row()
		{
		}

		public Row(Range rowCodeCell)
		{
			Id = rowCodeCell.Value.ToString();

			Code = rowCodeCell.Value.ToString();

			ElementTitle = CommonMethods.GetRowOrColumnTitle(rowCodeCell, true);

			Tag = SettingsManager.Settings.Tags.RowPrefix.Value +
				CommonMethods.GetTagFromMarkup(rowCodeCell, true) == "" ?
				CommonMethods.GetTag(Id) :
				CommonMethods.GetTagFromMarkup(rowCodeCell, true);
		}

		[XmlAttribute()]
		public string Id { get; set; }

		[XmlAttribute()]
		public string Code { get; set; }

		[XmlAttribute()]
		public string ElementTitle { get; set; }

		[XmlAttribute()]
		public string Tag { get; set; }
	}
}