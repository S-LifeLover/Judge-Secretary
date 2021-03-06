﻿using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace JudgeSecretary
{
	public class OrderParser
	{
		public OrderInfo Parse(params string[] text)
		{
			var dateAndCaseNumberRegex = new Regex(@"«(?<Day>\w+)» (?<Month>\w+) (?<Year>\d+)\s+года\s+производство\s+(?<CaseNumber>[\w+№а-я-\/]+)", RegexOptions.Singleline);
			var personInfoRegex = new Regex(@"должника\s*(?<FullName>[а-яА-ЯёЁ]+\s+[а-яА-ЯёЁ]+\s*[а-яА-ЯёЁ]*\s*[а-яА-ЯёЁ]*),*\s*(ИНН \d+)?,*\s+(?<BirthDate>\d+[\.\s]\d+[\.\s]\d+)\s+года\s+рождения,\s*(уроженца|уроженка) (?<BirthPlace>[\dа-яА-ЯёЁ\s\.\,\-;\/]+),\s*место\s*работы:?\s*(?<WorkPlace>[\dа-яА-ЯёЁ\s\.\,\-;\/]+),\s*(проживающего|проживающей)\s*по\s*адресу:?\s*(?<ResidencePlace>[\dа-яА-ЯёЁ\s\.\,\/\-;\/]+),?\s*(в\sпользу|задолженность - )");
			//var testRegex = new Regex(@"должнику\s*(?<FullName>[а-яА-ЯёЁ]+\s+[а-яА-ЯёЁ]+\s+[а-яА-ЯёЁ]+)");
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
					person.BirthPlace = personInfoMatch.Groups["BirthPlace"].Value;
					person.ResidencePlace = personInfoMatch.Groups["ResidencePlace"].Value;
					person.WorkPlace = personInfoMatch.Groups["WorkPlace"].Value;

					if (!persons.Contains(person))
					{
						persons.Add(person);
					}
				}
			}

			result.Persons = persons.ToArray();

			return result;
		}
	}
}
