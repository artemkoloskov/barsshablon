using System.Xml.Serialization;
using System.Xml.Schema;

namespace БАРСШаблон
{
    [System.Serializable()]
    [XmlType(AnonymousType = true)]
    public partial class Таблица
    {
        public Таблица()
        {
        }

        private СвободнаяЯчейка[] свободнаяЯчейка;
        private Строка[] строки;
        private Столбец[] столбцы;
        private string идентификатор;
        private string код;
        private string наименование;
        private string тег;
        private string ссылкаНаМетодическийСправочник;
        private bool ручноеДобавлениеСтрок = false;

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
        [XmlArrayItem("Строка", typeof(Строка), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
        public Строка[] Строки
        {
            get
            {
                return строки;
            }
            set
            {
                строки = value;
            }
        }

        [XmlArray(Form = XmlSchemaForm.Unqualified)]
        [XmlArrayItem("Столбец", typeof(Столбец), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
        public Столбец[] Столбцы
        {
            get
            {
                return столбцы;
            }
            set
            {
                столбцы = value;
            }
        }

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
        public string СсылкаНаМетодическийСправочник
        {
            get
            {
                return ссылкаНаМетодическийСправочник;
            }
            set
            {
                ссылкаНаМетодическийСправочник = value;
            }
        }

        [XmlAttribute()]
        public bool РучноеДобавлениеСтрок
        {
            get
            {
                return ручноеДобавлениеСтрок;
            }
            set
            {
                ручноеДобавлениеСтрок = value;
            }
        }
    }
}