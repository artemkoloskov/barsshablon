using System.Xml.Serialization;
using System.Xml.Schema;
using System.Collections.Generic;

namespace БАРСШаблон
{
	[System.Serializable()]
	[XmlType(AnonymousType = true)]
	[XmlRoot(Namespace = "", IsNullable = false)]
	public class ОписаниеФормы
	{
		public ОписаниеФормы()
		{
		}

		public ОписаниеФормы(Мета мета, List<Таблица> списокТаблиц, List<СвободнаяЯчейка> списокСвободныхЯчеек)
		{
			this.мета = мета;
			справочники = new Справочник[] { };
			структура = new Структура(списокТаблиц, списокСвободныхЯчеек);
		}

		private Мета мета;
		private Структура структура;
		private string меню = "";
		private Справочник[] справочники;

		[XmlElement("Мета", Form = XmlSchemaForm.Unqualified)]
		public Мета Мета
		{
			get => мета;
			set => мета = value;
		}

		[XmlElement("Структура", Form = XmlSchemaForm.Unqualified)]
		public Структура Структура
		{
			get => структура;
			set => структура = value;
		}

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Меню
		{
			get => меню;
			set => меню = value;
		}

		[XmlArray(Form = XmlSchemaForm.Unqualified)]
		[XmlArrayItem("Справочник", typeof(Справочник), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
		public Справочник[] Справочники
		{
			get => справочники;
			set => справочники = value;
		}
	}
}