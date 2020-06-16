using System.Xml.Serialization;
using System.Xml.Schema;
using System.Collections.Generic;

namespace БАРСШаблон
{
    [System.Serializable()]
    [XmlType(AnonymousType = true)]
    public partial class Структура
    {
        public Структура ()
        {
        }

        public Структура(List<Таблица> списокТаблиц, List<СвободнаяЯчейка> списокСвободныхЯчеек)
        {
            if (списокТаблиц.Count > 0)
            {
                таблицы = new Таблица[списокТаблиц.Count];

                int i = 0;

                foreach (Таблица таблица in списокТаблиц)
                {
                    таблицы[i] = таблица;
                }
            }

            if (списокСвободныхЯчеек.Count > 0)
            {
                свободныеЯчейки = new СвободнаяЯчейка[списокСвободныхЯчеек.Count];

                int i = 0;

                foreach (СвободнаяЯчейка cвободнаяЯчейка in списокСвободныхЯчеек)
                {
                    свободныеЯчейки[i] = cвободнаяЯчейка;
                }
            }
        }

        private СвободнаяЯчейка[] свободныеЯчейки;
        private Таблица[] таблицы;

        [XmlElement("СвободнаяЯчейка", Form = XmlSchemaForm.Unqualified)]
        public СвободнаяЯчейка[] СвободныеЯчейки
        {
            get
            {
                return свободныеЯчейки;
            }
            set
            {
                свободныеЯчейки = value;
            }
        }

        [XmlElement("Таблица", Form = XmlSchemaForm.Unqualified)]
        public Таблица[] Таблицы
        {
            get
            {
                return таблицы;
            }
            set
            {
                таблицы = value;
            }
        }
    }
}