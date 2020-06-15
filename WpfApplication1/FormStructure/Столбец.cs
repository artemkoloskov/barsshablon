using System.Configuration;
using System.Xml.Serialization;

namespace БАРСШаблон
{
    [System.Serializable()]
    [XmlType(AnonymousType = true)]
    public partial class Столбец
    {
        public Столбец()
        {
        }

        public Столбец(string кодСтолбца, string типСтолбца)
        {
            идентификатор = кодСтолбца;
            код = кодСтолбца;
            тип = типСтолбца;
            тег = ConfigurationManager.AppSettings.Get("СтолбецТегПрефикс") + CommonMethods.GetTagName(идентификатор);
            описание = CommonMethods.GetSerializedType(тип);
        }

        private string идентификатор;
        private string код;
        private string наименованиеЭлемента;
        private string тег;
        private string тип;
        private string описание;

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

        [XmlAttribute()]
        public string Тип
        {
            get
            {
                return тип;
            }
            set
            {
                тип = value;
            }
        }

        [XmlAttribute()]
        public string Описание
        {
            get
            {
                return описание;
            }
            set
            {
                описание = value;
            }
        }
    }
}