﻿using System.Xml.Serialization;

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

		public string ToXML()
		{
			using (var stringwriter = new System.IO.StringWriter())
			{
				var serializer = new XmlSerializer(this.GetType());
				serializer.Serialize(stringwriter, this);
				return stringwriter.ToString();
			}
		}
	}
}
