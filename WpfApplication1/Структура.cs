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
        private Таблица[] таблица;

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

        [XmlArray(Form = XmlSchemaForm.Unqualified)]
        [XmlArrayItem("Таблица", typeof(Таблица), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
        public Таблица[] Таблица
        {
            get
            {
                return таблица;
            }
            set
            {
                таблица = value;
            }
        }
    }
}