using System.Xml.Serialization;

namespace БАРСШаблон
{
    [System.Serializable()]
    [XmlType(AnonymousType = true)]
    public partial class Строка
    {
        public Строка()
        {
        }

        private string идентификатор;
        private string код;
        private string наименованиеЭлемента;
        private string тег;

        [XmlAttribute()]
        public string Идентификатор
        {
            get
            {
                return идентификатор;
            }
            set
            {
                идентификатор = value;
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
        public string НаименованиеЭлемента
        {
            get
            {
                return наименованиеЭлемента;
            }
            set
            {
                наименованиеЭлемента = value;
            }
        }

        [XmlAttribute()]
        public string Тег
        {
            get
            {
                return тег;
            }
            set
            {
                тег = value;
            }
        }
    }
}