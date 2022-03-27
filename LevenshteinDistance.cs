/* Класс содержит описание дистанции Дамерау-Левенштейна, определяющей
 * примерное равенство двух строк, которое измеряется собственно дистанцией.
 * (https://ru.wikipedia.org/wiki/%D0%A0%D0%B0%D1%81%D1%81%D1%82%D0%BE%D1%8F%D0%BD%D0%B8%D0%B5_%D0%94%D0%B0%D0%BC%D0%B5%D1%80%D0%B0%D1%83_%E2%80%94_%D0%9B%D0%B5%D0%B2%D0%B5%D0%BD%D1%88%D1%82%D0%B5%D0%B9%D0%BD%D0%B0)
 * Дистанция - это количество изменений которые нужно произвести с одной из
 * сравниваемых строк, для того, чтобы получить из нее вторую из сравниваемых
 * строк. Реализация дистанции взята с ресурса:
 * https://www.csharpstar.com/csharp-string-distance-algorithm/#:~:text=Damerau%2DLevenshtein%20Distance%20Algorithm%3A&text=The%20classical%20Levenshtein%20distance%20only,as%20the%20Damerau%E2%80%93Levenshtein%20distance.
*/

using System;

namespace БАРСШаблон
{
	/// <summary>
	/// Содержит приблзительное сравнение строк
	/// </summary>
	public class DamerauLevenshteinDistance
	{
		/// <summary>
		/// Подсчитать дистанцию между двумя строками
		/// </summary>
		public static int DistanceBetweenStrings(string s, string t)
		{
			var bounds = new { Height = s.Length + 1, Width = t.Length + 1 };

			int[,] matrix = new int[bounds.Height, bounds.Width];

			for (int height = 0; height < bounds.Height; height++) { matrix[height, 0] = height; };
			for (int width = 0; width < bounds.Width; width++) { matrix[0, width] = width; };

			for (int height = 1; height < bounds.Height; height++)
			{
				for (int width = 1; width < bounds.Width; width++)
				{
					int cost = (s[height - 1] == t[width - 1]) ? 0 : 1;
					int insertion = matrix[height, width - 1] + 1;
					int deletion = matrix[height - 1, width] + 1;
					int substitution = matrix[height - 1, width - 1] + cost;

					int distance = Math.Min(insertion, Math.Min(deletion, substitution));

					if (height > 1 && width > 1 && s[height - 1] == t[width - 2] && s[height - 2] == t[width - 1])
					{
						distance = Math.Min(distance, matrix[height - 2, width - 2] + cost);
					}

					matrix[height, width] = distance;
				}
			}

			return matrix[bounds.Height - 1, bounds.Width - 1];
		}
	}
}
