using System.Xml.Serialization;

namespace БАРСШаблон
{
	[System.Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class ЭлементСправочника
	{
		public ЭлементСправочника()
		{
		}

		private string код;
		private string наименование;

		[XmlAttribute()]
		public string Код
		{
			get => код;
			set => код = value;
		}

		[XmlAttribute()]
		public string Наименование
		{
			get => наименование;
			set => наименование = value;
		}
	}
}