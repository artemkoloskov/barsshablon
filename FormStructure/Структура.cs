using System.Collections.Generic;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace БАРСШаблон
{
	[System.Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class Структура
	{
		public Структура()
		{
		}

		public Структура(List<Таблица> списокТаблиц, List<СвободнаяЯчейка> списокСвободныхЯчеек)
		{
			if (списокТаблиц.Count > 0)
			{
				Таблицы = списокТаблиц.ToArray();
			}

			if (списокСвободныхЯчеек.Count > 0)
			{
				СвободныеЯчейки = списокСвободныхЯчеек.ToArray();
			}
		}

		private СвободнаяЯчейка[] свободныеЯчейки;
		private Таблица[] таблицы;

		[XmlElement("СвободнаяЯчейка", Form = XmlSchemaForm.Unqualified)]
		public СвободнаяЯчейка[] СвободныеЯчейки { get => свободныеЯчейки; set => свободныеЯчейки = value; }

		[XmlElement("Таблица", Form = XmlSchemaForm.Unqualified)]
		public Таблица[] Таблицы { get => таблицы; set => таблицы = value; }
	}
}