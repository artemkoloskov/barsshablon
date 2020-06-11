namespace БАРСШаблон.DataTypes
{
    public class Финансовый
    {
        public int Точность = 2;
        public string ValueRange = "";
        public bool ОбязательноДляЗаполнения = false;
        public bool ТолькоЧтение = false;
        public string Комментарий = "";
        public bool ЯвляетсяКлючевым = false;
        public string ЗначениеПоУмолчанию = "";
        public string ДействиеСПолем = "Суммировать";
    }
}
