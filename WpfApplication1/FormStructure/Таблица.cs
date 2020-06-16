using System.Xml.Serialization;
using System.Xml.Schema;
using System.Configuration;
using Microsoft.Office.Interop.Excel;
using System;
using БАРСШаблон.DataTypes;

namespace БАРСШаблон
{
    [System.Serializable()]
    [XmlType(AnonymousType = true)]
    public partial class Таблица
    {
        public Таблица()
        {
        }

        public Таблица(string кодТаблицы)
        {
            идентификатор = кодТаблицы;
            код = кодТаблицы;
            тег = ConfigurationManager.AppSettings.Get("ТаблицаТегПрефикс") + CommonMethods.ПолчитьТег(идентификатор);
        }

        private СвободнаяЯчейка[] свободныеЯчейки;
        private Строка[] строки;
        private Столбец[] столбцы;
        private string идентификатор;
        private string код;
        private string наименование;
        private string тег;
        private string ссылкаНаМетодическийСправочник;
        private bool ручноеДобавлениеСтрок = false;

        [XmlArray(Form = XmlSchemaForm.Unqualified)]
        [XmlArrayItem("СвободнаяЯчейка", typeof(СвободнаяЯчейка), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
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

        internal static Таблица ПолучитьТаблицуИз(Worksheet sheet)
        {
            Столбец столбец1 = new Столбец("1", typeof(Целочисленный).ToString().Split('.')[2]);
            Столбец столбец2 = new Столбец("2", typeof(Финансовый).ToString().Split('.')[2]);

            Строка строка1 = new Строка() { Идентификатор = "001", Код = "001", НаименованиеЭлемента = "Охуеть", Тег = "Охт" };
            Строка строка2 = new Строка() { Идентификатор = "002", Код = "002", НаименованиеЭлемента = "Заебись", Тег = "Збс" };

            Таблица таблица1 = new Таблица()
            {
                Идентификатор = "Таблица1",
                Код = "Тбл1",
                Наименование = "Крутая ваще таблица",
                РучноеДобавлениеСтрок = false,
                Тег = "КртВщТабла",
                Столбцы = new Столбец[] { столбец1, столбец2 },
                Строки = new Строка[] { строка1, строка2 },
                СвободныеЯчейки = new СвободнаяЯчейка[] { new СвободнаяЯчейка("Суки", typeof(Целочисленный).ToString().Split('.')[2]) },
            };

            throw new NotImplementedException();
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