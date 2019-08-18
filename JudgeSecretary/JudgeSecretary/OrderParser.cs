using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace JudgeSecretary
{
	public class OrderParser
	{
		public OrderInfo Parse(string[] text)
		{
			var dateAndCaseNumberRegex = new Regex(@"«(?<Day>\w+)» (?<Month>\w+) (?<Year>\d+)\s+года\s+производство\s+(?<CaseNumber>[\w+\W]+)");
			var personInfoRegex = new Regex(@"(?<FullName>[а-яА-Я]+\s+[а-яА-Я]+\s+[а-яА-Я]+)\s+(?<BirthDate>\d+\.\d+\.\d+)\s+года\s+рождения");
			var result = new OrderInfo();
			var persons = new List<OrderInfo.PersonInfo>();
			foreach (var line in text)
			{
				var trimmedLine = line.Trim();

				var dateAndCaseNumberMatch = dateAndCaseNumberRegex.Match(trimmedLine);
				if (dateAndCaseNumberMatch.Success)
				{
					result.Day = dateAndCaseNumberMatch.Groups["Day"].Value.Replace("I", "1");
					result.Month = dateAndCaseNumberMatch.Groups["Month"].Value;
					result.Year = dateAndCaseNumberMatch.Groups["Year"].Value;
					result.CaseNumber = dateAndCaseNumberMatch.Groups["CaseNumber"].Value;
				}

				foreach (Match personInfoMatch in personInfoRegex.Matches(trimmedLine))
				{
					var person = new OrderInfo.PersonInfo();
					person.FullName = personInfoMatch.Groups["FullName"].Value;
					person.BirthDate = personInfoMatch.Groups["BirthDate"].Value;
					persons.Add(person);
				}
			}

			result.Persons = persons.ToArray();

			return result;
		}
	}
}
