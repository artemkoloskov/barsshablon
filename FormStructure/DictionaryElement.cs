using System.Xml.Serialization;

namespace БАРСШаблон
{
	[System.Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class DictionaryElement
	{
		public DictionaryElement()
		{
		}

		[XmlAttribute("Код")]
		public string Code { get; set; }

		[XmlAttribute()]
		public string Title { get; set; }
	}
}