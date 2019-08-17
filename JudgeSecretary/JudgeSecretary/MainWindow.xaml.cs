using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Forms;
using Xceed.Words.NET;
using MessageBox = System.Windows.Forms.MessageBox;
using Path = System.IO.Path;

namespace JudgeSecretary
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();
		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			using (var fbd = new FolderBrowserDialog())
			{
				fbd.Description = "Выберете папку с приказами";
				var result = fbd.ShowDialog();

				if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
				{
					var files = Directory.GetFiles(fbd.SelectedPath);

					fbd.Description = "Выберете папку, куда сохранить исполнительные";
					result = fbd.ShowDialog();

					if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
					{
						var destinationFolder = fbd.SelectedPath;

						foreach (var file in files)
						{
							var destinationFilePath = Path.Combine(destinationFolder, Path.ChangeExtension(Path.GetFileName(file), "txt"));

							/*
							var arguments = $"\"{file}\" /lang Mixed /out \"{destinationFilePath}\" /quit";

							System.Diagnostics.Process process1 = new System.Diagnostics.Process();
							process1.StartInfo.FileName = FineReaderPathTextBox.Text;
							process1.StartInfo.Arguments = arguments;
							process1.Start();
							process1.WaitForExit();
							process1.Close();
							*/

							var content = File.ReadAllLines(destinationFilePath);
							var dateAndCaseNumberRegex = new Regex(@"«(?<Day>\w+)» (?<Month>\w+) (?<Year>\d+)\s+года\s+производство\s+(?<CaseNumber>[\w+\W]+)");
							var manInfoRegex = new Regex(@"Взыскать\s*(солидарно)?\s+с\s*(?<FullName>[а-яА-Я]+\s+[а-яА-Я]+\s+[а-яА-Я]+)\s+(?<BirthDate>\d+\.\d+\.\d+)");
							string day = "xxxx";
							string month = "xxxx";
							string year = "xxxx";
							string caseNumber = "xxxx";
							string fullName = "xxxx";
							string birthDate = "xxxx";
							foreach (var line in content)
							{
								var match = dateAndCaseNumberRegex.Match(line);
								if (match.Success)
								{
									day = match.Groups["Day"].Value.Replace("I", "1");
									month = match.Groups["Month"].Value;
									year = match.Groups["Year"].Value;
									caseNumber = match.Groups["CaseNumber"].Value;
								}
								var match2 = manInfoRegex.Match(line);
								if (match2.Success)
								{
									fullName = match2.Groups["FullName"].Value;
									birthDate = match2.Groups["BirthDate"].Value;
								}
							}

							var docxFilePath = Path.ChangeExtension(destinationFilePath, "docx");
							File.Copy("Template.docx", docxFilePath, true);

							using (var document = DocX.Load(docxFilePath))
							{
								document.ReplaceText("{CaseNumber}", caseNumber);
								document.ReplaceText("{Day}", day);
								document.ReplaceText("{Month}", month);
								document.ReplaceText("{Year}", year.Substring(year.Length - 2, 2));

								document.ReplaceText("{FullName}", fullName);
								document.ReplaceText("{BirthDate}", birthDate);

								document.Save();
							}
						}

						MessageBox.Show("Готово");
					}
				}
			}
		}
	}
}
