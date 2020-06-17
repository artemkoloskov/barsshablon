using System.Xml.Serialization;
using System.Xml.Schema;
using System;
using System.Configuration;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace БАРСШаблон
{
    [System.Serializable()]
    [XmlType(AnonymousType = true)]
    public partial class Мета
    {
        public Мета()
        {
        }

        public Мета(Workbook workbook)
        {
            наименование = ПолучитьНаименованиеИз(workbook.Sheets[1]);
            идентификатор = идентификатор + CommonMethods.ПолчитьТег(наименование);
            группа = группа + DateTime.Today.Year;
            датаНачалаДействия = датаНачалаДействия.Replace("0001", DateTime.Today.Year.ToString());
            датаОкончанияДействия = датаОкончанияДействия.Replace("9999", DateTime.Today.Year.ToString());
            датаПоследнегоИзменения = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");
            тег = идентификатор;
        }

        private static string ПолучитьНаименованиеИз(Worksheet sheet)
        {
            Dictionary<string, double> возможныеНаименования = new Dictionary<string, double>();

            Range usedRange = (Range)sheet.UsedRange;

            foreach (Range column in usedRange.Columns)
            {
                Range topCellInColumn = НайтиВКолонкеВерхнююНеПустуюЯчейку(column);

                if(!возможныеНаименования.ContainsKey(topCellInColumn.Value.ToString()))
                {
                    возможныеНаименования.Add(topCellInColumn.Value.ToString(), ПолучитьВероятностьТогоЧтоВЯчейкеНаименование(topCellInColumn));
                }
            }

            KeyValuePair<string, double> наиболееВероятноеНаименование = new KeyValuePair<string, double>("", 0);

            foreach (var возможноеНаименование in возможныеНаименования)
            {
                if (возможноеНаименование.Value > наиболееВероятноеНаименование.Value)
                {
                    наиболееВероятноеНаименование = возможноеНаименование;
                }
            }

            return наиболееВероятноеНаименование.Key;
        }

        private static Range НайтиВКолонкеВерхнююНеПустуюЯчейку(Range column)
        {
            foreach (Range cell in column.Cells)
            {
                if (cell.Value != null && cell.Value.ToString() != "" && cell.Value.ToString() != " ")
                {
                    return cell;
                }
            }

            return null;
        }

        private static double ПолучитьВероятностьТогоЧтоВЯчейкеНаименование(Range cell)
        {
            double вероятность = 0;

            double весДлины = double.Parse(ConfigurationManager.AppSettings.Get("МетаВесДлиныПотенциальногоНаименования"));
            double весНомераСтроки = double.Parse(ConfigurationManager.AppSettings.Get("МетаВесНомераСтрокиПотенциальногоНаименования"));
            double весНомераСтолбца = double.Parse(ConfigurationManager.AppSettings.Get("МетаВесНомераСтолбцаПотенциальногоНаименования"));
            double весКолЯчеекВОбъедЯчейке = double.Parse(ConfigurationManager.AppSettings.Get("МетаВесКоличестваЯчеекВОбъединеннойЯчейкеПотенциальногоНаименования"));
            double весГраницыВнизу = double.Parse(ConfigurationManager.AppSettings.Get("МетаВесГраницыВнизуПотенциальногоНаименования"));
            double весГраницыВверху = double.Parse(ConfigurationManager.AppSettings.Get("МетаВесГраницыВверхуПотенциальногоНаименования"));
            double весГраницыСлева = double.Parse(ConfigurationManager.AppSettings.Get("МетаВесГраницыСлеваПотенциальногоНаименования"));
            double весГраницыСправа = double.Parse(ConfigurationManager.AppSettings.Get("МетаВесГраницыСправаПотенциальногоНаименования"));
            double весВыравнПоСередине = double.Parse(ConfigurationManager.AppSettings.Get("МетаВесВыравниванияПоСерединеПотенциальногоНаименования"));
            double весВыравнСлева = double.Parse(ConfigurationManager.AppSettings.Get("МетаВесВыравниванияСлеваПотенциальногоНаименования"));
            double весВыравнСправа = double.Parse(ConfigurationManager.AppSettings.Get("МетаВесВыравниванияСправаПотенциальногоНаименования"));
            double весЖирностиТекста = double.Parse(ConfigurationManager.AppSettings.Get("МетаВесЖирностиТекстаПотенциальногоНаименования"));
            double весПустойСтроки = double.Parse(ConfigurationManager.AppSettings.Get("МетаВесПустойСтрокиПодЯчейкойПотенциальногоНаименования"));

            вероятность += cell.Value.ToString().Length * весДлины;
            вероятность += cell.Row * весНомераСтроки;
            вероятность += cell.Column * весНомераСтолбца;
            вероятность += ПолучитьКоличествоЯчеекВОбъединении(cell) * весКолЯчеекВОбъедЯчейке;
            вероятность += cell.Borders[XlBordersIndex.xlEdgeBottom].LineStyle == (int)XlLineStyle.xlLineStyleNone ? 0 : 1 * весГраницыВнизу;
            вероятность += cell.Borders[XlBordersIndex.xlEdgeTop].LineStyle == (int)XlLineStyle.xlLineStyleNone ? 0 : 1 * весГраницыВверху;
            вероятность += cell.Borders[XlBordersIndex.xlEdgeLeft].LineStyle == (int)XlLineStyle.xlLineStyleNone ? 0 : 1 * весГраницыСлева;
            вероятность += cell.Borders[XlBordersIndex.xlEdgeRight].LineStyle == (int)XlLineStyle.xlLineStyleNone ? 0 : 1 * весГраницыСправа;
            вероятность += cell.HorizontalAlignment == (int)XlHAlign.xlHAlignCenter ? 1 : 0 * весВыравнПоСередине;
            вероятность += cell.HorizontalAlignment == (int)XlHAlign.xlHAlignLeft ? 1 : 0 * весВыравнСлева;
            вероятность += cell.HorizontalAlignment == (int)XlHAlign.xlHAlignRight ? 1 : 0 * весВыравнСправа;
            вероятность += cell.Font.Bold ? 1 : 0 * весЖирностиТекста;
            вероятность += ПолучитьКоличествоПустыхСтрокПодЯчейкой(cell) * весПустойСтроки;

            return вероятность;
        }

        private static int ПолучитьКоличествоЯчеекВОбъединении(Range cell)
        {
            int количество = 0;

            if (cell.MergeCells)
            {
                for (int i = 1; i < 10; i++)
                {
                    if (cell.Offset[i, 0].MergeCells)
                    {
                        for (int j = 1; j < 10; j++)
                        {
                            if (cell.Offset[i, j].MergeCells)
                            {
                                количество++;
                            }
                            else
                            {
                                return количество;
                            }
                        }
                    }
                    else
                    {
                        return количество;
                    }
                }
            }

            return количество;
        }

        private static int ПолучитьКоличествоПустыхСтрокПодЯчейкой(Range cell)
        {
            int количество = 0;

            do
            {
                количество++;
            } while (количество < 10 && (cell.Offset[количество, 0].Value == null || cell.Offset[количество, 0].Value.ToString() == "" ||
                    cell.Offset[количество, 0].Value.ToString() == " "));

            return количество;
        }

        private string версияМетаописания = ConfigurationManager.AppSettings.Get("МетаВерсияМетаописания");
        private string идентификатор = ConfigurationManager.AppSettings.Get("МетаИдентификатор");
        private string наименование = "";
        private string группа = ConfigurationManager.AppSettings.Get("МетаГруппа");
        private string датаНачалаДействия = ConfigurationManager.AppSettings.Get("МетаДатаНачалаДействия");
        private string датаОкончанияДействия = ConfigurationManager.AppSettings.Get("МетаДатаОкончанияДействия");
        private string авторство = ConfigurationManager.AppSettings.Get("МетаАвторство");
        private string датаПоследнегоИзменения = "";
        private string номерВерсии = ConfigurationManager.AppSettings.Get("МетаНомерВерсии");
        private string расположениеШапки = ConfigurationManager.AppSettings.Get("МетаРасположениеШапки");
        private string хост = Environment.MachineName;
        private string ссылкаНаМетодическийСправочник = "";
        private string ссылкаНаВнешнююСправку = "";
        private string версияФорматаМетаструктуры = ConfigurationManager.AppSettings.Get("МетаВерсияФорматаМетаструктуры");
        private string тег = "";

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string ВерсияМетаописания
        {
            get
            {
                return версияМетаописания;
            }
            set
            {
                версияМетаописания = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
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

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
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

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string Группа
        {
            get
            {
                return группа;
            }
            set
            {
                группа = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string ДатаНачалаДействия
        {
            get
            {
                return датаНачалаДействия;
            }
            set
            {
                датаНачалаДействия = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string ДатаОкончанияДействия
        {
            get
            {
                return датаОкончанияДействия;
            }
            set
            {
                датаОкончанияДействия = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string Авторство
        {
            get
            {
                return авторство;
            }
            set
            {
                авторство = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string ДатаПоследнегоИзменения
        {
            get
            {
                return датаПоследнегоИзменения;
            }
            set
            {
                датаПоследнегоИзменения = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string НомерВерсии
        {
            get
            {
                return номерВерсии;
            }
            set
            {
                номерВерсии = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string РасположениеШапки
        {
            get
            {
                return расположениеШапки;
            }
            set
            {
                расположениеШапки = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string Хост
        {
            get
            {
                return хост;
            }
            set
            {
                хост = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
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

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string СсылкаНаВнешнююСправку
        {
            get
            {
                return ссылкаНаВнешнююСправку;
            }
            set
            {
                ссылкаНаВнешнююСправку = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string ВерсияФорматаМетаструктуры
        {
            get
            {
                return версияФорматаМетаструктуры;
            }
            set
            {
                версияФорматаМетаструктуры = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
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