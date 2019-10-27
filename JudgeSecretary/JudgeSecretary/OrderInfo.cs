using System;

namespace JudgeSecretary
{
	public class OrderInfo
	{
		public string Day { get; set; }
		public string Month { get; set; }
		public string Year { get; set; }
		public string CaseNumber { get; set; }
		public PersonInfo[] Persons { get; set; }
		public string StateDuty { get; set; }

		public class PersonInfo
		{
			public string FullName { get; set; }
			public string BirthDate { get; set; }
			public string BirthPlace { get; set; }
			public string ResidencePlace { get; set; }
			public string WorkPlace { get; set; }

			public override bool Equals(object obj)
			{
				var info = obj as PersonInfo;
				return info != null &&
					   FullName == info.FullName &&
					   BirthDate == info.BirthDate &&
					   BirthPlace == info.BirthPlace &&
					   ResidencePlace == info.ResidencePlace &&
					   WorkPlace == info.WorkPlace;
			}
		}
	}
}
