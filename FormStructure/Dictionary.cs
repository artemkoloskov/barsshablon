using System.Xml.Schema;
using System.Xml.Serialization;

namespace БАРСШаблон
{
	[System.Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class Dictionary
	{
		public Dictionary()
		{
		}

		[XmlArray(Form = XmlSchemaForm.Unqualified)]
		[XmlArrayItem("ЭлементСправочника", typeof(DictionaryElement), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
		public DictionaryElement[] Elements { get; set; }

		[XmlAttribute("Код")]
		public string Code { get; set; }

		[XmlAttribute("Наименование")]
		public string Title { get; set; }
	}
}