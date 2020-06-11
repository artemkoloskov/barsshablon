namespace БАРСШаблон.DataTypes
{
    public class Учреждение
    {
        public bool ОбязательноДляЗаполнения = false;
        public bool ТолькоЧтение = false;
        public string Комментарий = "";
        public bool ЯвляетсяКлючевым = true;
        public string ЗначениеПоУмолчанию = "";
        public string ДействиеСПолем = "БезИтогов";
    }
}
