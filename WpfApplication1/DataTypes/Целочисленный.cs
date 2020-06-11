using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace БАРСШаблон.DataTypes
{
    public class Целочисленный
    {
        public int Точность = 0;
        public string ValueRange = "";
        public bool ОбязательноДляЗаполнения = false;
        public bool ТолькоЧтение = false;
        public string Комментарий = "";
        public bool ЯвляетсяКлючевым = false;
        public string ЗначениеПоУмолчанию = "";
        public string ДействиеСПолем = "Суммировать";
    }
}
