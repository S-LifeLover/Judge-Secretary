﻿using System.IO;
using System.Windows;
using System.Windows.Forms;
using Xceed.Words.NET;
using MessageBox = System.Windows.Forms.MessageBox;
using Path = System.IO.Path;
using Word = Microsoft.Office.Interop.Word;

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

		private void OldLogicButton_Click(object sender, RoutedEventArgs e)
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

							var arguments = $"\"{file}\" /lang Mixed /out \"{destinationFilePath}\" /quit";

							System.Diagnostics.Process process1 = new System.Diagnostics.Process();
							process1.StartInfo.FileName = FineReaderPathTextBox.Text;
							process1.StartInfo.Arguments = arguments;
							process1.Start();
							process1.WaitForExit();
							process1.Close();

							var content = File.ReadAllLines(destinationFilePath);
							var parser = new OrderParser();
							var orderInfo = parser.Parse(content);

							var docxFilePath = Path.ChangeExtension(destinationFilePath, "docx");
							File.Copy("Template.docx", docxFilePath, true);

							using (var document = DocX.Load(docxFilePath))
							{
								document.ReplaceText("{CaseNumber}", orderInfo.CaseNumber);
								document.ReplaceText("{Day}", orderInfo.Day);
								document.ReplaceText("{Month}", orderInfo.Month);
								document.ReplaceText("{Year}", orderInfo.Year.Substring(orderInfo.Year.Length - 2, 2));

								document.ReplaceText("{FullName}", orderInfo.Persons[0].FullName);
								document.ReplaceText("{BirthDate}", orderInfo.Persons[0].BirthDate);

								document.Save();
							}
						}

						MessageBox.Show("Готово");
					}
				}
			}
		}

		private void TaxButton_Click(object sender, RoutedEventArgs e)
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

					if (result == System.Windows.Forms.DialogResult.OK &&
						!string.IsNullOrWhiteSpace(fbd.SelectedPath))
					{
						var destinationFolder = fbd.SelectedPath;

						foreach (var file in files)
						{
							var docxFilePath = Path.Combine(destinationFolder,
								Path.ChangeExtension(Path.GetFileName(file), "docx"));

							string content = string.Empty;

							var application = new Word.Application();
							Word._Document document = application.Documents.Open(file);
							content = document.Content.Text;
							document.Close();
							application.Quit();

							var parser = new OrderParser();
							var orderInfo = parser.Parse(content);

							File.Copy("Template.docx", docxFilePath, true);

							using (var docxDocument = DocX.Load(docxFilePath))
							{
								docxDocument.ReplaceText("{CaseNumber}", orderInfo.CaseNumber);
								docxDocument.ReplaceText("{Day}", orderInfo.Day);
								docxDocument.ReplaceText("{Month}", orderInfo.Month);
								docxDocument.ReplaceText("{Year}",
									orderInfo.Year.Substring(orderInfo.Year.Length - 2, 2));

								docxDocument.ReplaceText("{FullName}", orderInfo.Persons[0].FullName);
								docxDocument.ReplaceText("{BirthDate}", orderInfo.Persons[0].BirthDate);
								docxDocument.ReplaceText("{BirthPlace}", orderInfo.Persons[0].BirthPlace);
								docxDocument.ReplaceText("{ResidencePlace}", orderInfo.Persons[0].ResidencePlace);

								docxDocument.Save();
							}
						}

						MessageBox.Show("Готово");
					}
				}
			}
		}
	}
}
