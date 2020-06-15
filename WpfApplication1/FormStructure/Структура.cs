using System.Xml.Serialization;
using System.Xml.Schema;

namespace БАРСШаблон
{
    [System.Serializable()]
    [XmlType(AnonymousType = true)]
    public partial class Структура
    {
        public Структура ()
        {
        }

        private СвободнаяЯчейка[] свободнаяЯчейка;
        private Таблица[] таблицы;

        [XmlElement("СвободнаяЯчейка", Form = XmlSchemaForm.Unqualified)]
        public СвободнаяЯчейка[] СвободнаяЯчейка
        {
            get
            {
                return свободнаяЯчейка;
            }
            set
            {
                свободнаяЯчейка = value;
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