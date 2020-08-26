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

		[XmlAttribute()]
		public string Код { get; set; }

		[XmlAttribute()]
		public string Наименование { get; set; }
	}
}