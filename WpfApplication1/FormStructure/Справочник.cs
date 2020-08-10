using System.Xml.Serialization;
using System.Xml.Schema;

namespace БАРСШаблон
{
	[System.Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class Справочник
	{
		public Справочник()
		{
		}

		private ЭлементСправочника[] элементы;
		private string код;
		private string наименование;

		[XmlArray(Form = XmlSchemaForm.Unqualified)]
		[XmlArrayItem("ЭлементСправочника", typeof(ЭлементСправочника), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
		public ЭлементСправочника[] Элементы
		{
			get
			{
				return элементы;
			}
			set
			{
				элементы = value;
			}
		}

		[XmlAttribute()]
		public string Код
		{
			get
			{
				return код;
			}
			set
			{
				код = value;
			}
		}

		[XmlAttribute()]
		public string Наименование
		{
			get
			{
				return наименование;
			}
			set
			{
				наименование = value;
			}
		}
	}
}