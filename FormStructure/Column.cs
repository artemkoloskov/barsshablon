using Microsoft.Office.Interop.Excel;
using System;
using System.Xml.Serialization;

namespace БАРСШаблон
{
	[Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class Column
	{
		public Column()
		{
		}

		public Column(Range columnCodeCell, bool isKey)
		{
			Console.WriteLine(DateTime.Now + ": столбец " + columnCodeCell.Value.ToString() + ", начат");

			Id = columnCodeCell.Value.ToString();

			Code = columnCodeCell.Value.ToString();

			НаименованиеЭлемента = CommonMethods.GetRowOrColumnTitle(codesRangeCell: columnCodeCell, searchingForRow: false);

			Tag = SettingsManager.Settings.Tags.ColumnPrefix.Value +
				CommonMethods.GetTagFromMarkup(columnCodeCell, false) == "" ?
				CommonMethods.GetTag(Id) :
				CommonMethods.GetTagFromMarkup(columnCodeCell, false);

			IsKey = isKey;

			ColumnType = GetColumnType(columnCodeCell);

			Type = ColumnType.GetType().Name;

			Описание = CommonMethods.GetSerializedType(ColumnType);

			Console.WriteLine(DateTime.Now + ": столбец " + columnCodeCell.Value.ToString() + ", закончен");
		}

		private object GetColumnType(Range columnCodeCell)
		{
			return CommonMethods.GetCellType(columnCodeCell.Offset[1, 0].NumberFormat, IsKey);
		}

		[XmlAttribute()]
		public string Id { get; set; }

		[XmlAttribute()]
		public string Code { get; set; }

		[XmlAttribute()]
		public string НаименованиеЭлемента { get; set; }

		[XmlAttribute()]
		public string Tag { get; set; }

		[XmlAttribute()]
		public string Type { get; set; }

		[XmlAttribute()]
		public string Описание { get; set; }

		[XmlIgnore]
		public object ColumnType { get; }

		[XmlIgnore]
		public bool IsKey { get; }
	}
}