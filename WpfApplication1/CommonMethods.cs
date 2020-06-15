using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using БАРСШаблон.DataTypes;

namespace БАРСШаблон
{
    public static class CommonMethods
    {
        /// <summary>
        /// Сокращает строку то приемлемого полю тег вида
        /// </summary>
        /// <param name="идентификатор"></param>
        /// <returns></returns>
        public static string GetTagName(string идентификатор)
        {
            return идентификатор; //TODO
        }

        public static string GetSerializedType(string тип)
        {
            switch (тип)
            {
                case "ДатаВремя":
                    return new ДатаВремя().ToXML();
                case "Логический":
                    return new Логический().ToXML();
                case "Строковый":
                    return new Строковый().ToXML();
                case "Учреждение":
                    return new Учреждение().ToXML();
                case "Финансовый":
                    return new Финансовый().ToXML();
                case "Целочисленный":
                    return new Целочисленный().ToXML();
                default:
                    return "";
            }
        }
    }
}
