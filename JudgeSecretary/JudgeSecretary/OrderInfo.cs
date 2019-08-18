namespace JudgeSecretary
{
	public class OrderInfo
	{
		public string Day { get; set; }
		public string Month { get; set; }
		public string Year { get; set; }
		public string CaseNumber { get; set; }
		public PersonInfo[] Persons { get; set; }

		public class PersonInfo
		{
			public string FullName { get; set; }
			public string BirthDate { get; set; }
		}
	}
}
