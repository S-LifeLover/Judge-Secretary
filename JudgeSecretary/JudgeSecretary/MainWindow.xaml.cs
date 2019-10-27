using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Forms;
using Xceed.Words.NET;
using MessageBox = System.Windows.Forms.MessageBox;
using Path = System.IO.Path;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;

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
								/*
								docxDocument.ReplaceText("{CaseNumber}", orderInfo.CaseNumber ?? String.Empty);
								docxDocument.ReplaceText("{Day}", orderInfo.Day ?? String.Empty);
								docxDocument.ReplaceText("{Month}", orderInfo.Month ?? String.Empty);
								docxDocument.ReplaceText("{Year}",
									orderInfo.Year?.Substring(orderInfo.Year.Length - 2, 2) ?? String.Empty);

								docxDocument.ReplaceText("{FullName}", orderInfo.Persons[0].FullName ?? String.Empty);
								docxDocument.ReplaceText("{FullNameNominative}", orderInfo.Persons[0].FullName ?? String.Empty);
								docxDocument.ReplaceText("{BirthDate}", orderInfo.Persons[0].BirthDate ?? String.Empty);
								docxDocument.ReplaceText("{BirthPlace}", orderInfo.Persons[0].BirthPlace ?? String.Empty);
								docxDocument.ReplaceText("{ResidencePlace}", orderInfo.Persons[0].ResidencePlace ?? String.Empty);

								docxDocument.Save();
								*/
							}
						}

						MessageBox.Show("Готово");
					}
				}
			}
		}

		private void DataFileButton_OnClick(object sender, RoutedEventArgs e)
		{
			using (var ofd = new OpenFileDialog())
			{
				ofd.Title = "Выберете файл с данными";

				var result = ofd.ShowDialog();

				if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(ofd.FileName))
				{
					using (var fbd = new FolderBrowserDialog())
					{
						fbd.Description = "Выберете папку, куда сохранить исполнительные";
						result = fbd.ShowDialog();

						if (result == System.Windows.Forms.DialogResult.OK &&
							!string.IsNullOrWhiteSpace(fbd.SelectedPath))
						{
							var excel = new Microsoft.Office.Interop.Excel.Application();
							try
							{
								Workbook wb = null;
								var orders = new List<OrderInfo>();
								try
								{
									wb = excel.Workbooks.Open(ofd.FileName, ReadOnly: true);
									Worksheet sheet = wb.Sheets["Данные"];

									var consecutiveBlankLines = 0;

									for (var currentRowIndex = 2; currentRowIndex <= sheet.Rows.Count && consecutiveBlankLines < 100; currentRowIndex++)
									{
										var caseNumber = sheet.Rows.Cells[currentRowIndex, 1].Value as string;

										if (string.IsNullOrEmpty(caseNumber))
										{
											consecutiveBlankLines++;
											continue;
										}

										consecutiveBlankLines = 0;

										var order = new OrderInfo
										{
											CaseNumber = caseNumber
										};

										if (string.IsNullOrEmpty(order.CaseNumber))
										{
											break;
										}

										order.Persons = new[]
										{
											new OrderInfo.PersonInfo
											{
												FullName = sheet.Rows.Cells[currentRowIndex, 3].Value as string,
												ResidencePlace = sheet.Rows.Cells[currentRowIndex, 7].Value as string,
												BirthDate = sheet.Rows.Cells[currentRowIndex, 9].Value as string,
												BirthPlace = sheet.Rows.Cells[currentRowIndex, 10].Value as string
											}
										};

										var date = (DateTime)sheet.Rows.Cells[currentRowIndex, 5].Value;

										order.Day = date.Day.ToString();
										order.Month = date.ToString("MMMM", CultureInfo.GetCultureInfo("ru-ru"));
										order.Year = date.Year.ToString();

										order.StateDuty = sheet.Rows.Cells[currentRowIndex, 17].Value as string;

										orders.Add(order);
									}
								}
								catch (Exception)
								{
									wb?.Close(false);
									throw;
								}

								var destinationFolder = fbd.SelectedPath;

								foreach (var orderInfo in orders)
								{
									var docxFilePath = Path.Combine(destinationFolder,
										Path.ChangeExtension(MakeValidFileName(orderInfo.CaseNumber) + "_" + orderInfo.Persons.First().FullName.Replace(" ", "_"), "docx"));

									File.Copy("Template.docx", docxFilePath, true);

									using (var docxDocument = DocX.Load(docxFilePath))
									{
										docxDocument.ReplaceText("{CaseNumber}", orderInfo.CaseNumber ?? string.Empty);
										docxDocument.ReplaceText("{StateDuty}", orderInfo.StateDuty ?? string.Empty);
										docxDocument.ReplaceText("{Day}", orderInfo.Day);
										docxDocument.ReplaceText("{Month}", orderInfo.Month);
										docxDocument.ReplaceText("{Year}", orderInfo.Year?.Substring(orderInfo.Year.Length - 2, 2));
										docxDocument.ReplaceText("{FullYear}", orderInfo.Year);

										docxDocument.ReplaceText("{FullName}", orderInfo.Persons[0].FullName ?? string.Empty);
										docxDocument.ReplaceText("{FullNameNominative}", orderInfo.Persons[0].FullName ?? string.Empty);
										docxDocument.ReplaceText("{BirthDate}", orderInfo.Persons[0].BirthDate ?? string.Empty);
										docxDocument.ReplaceText("{BirthPlace}", orderInfo.Persons[0].BirthPlace ?? string.Empty);
										docxDocument.ReplaceText("{ResidencePlace}", orderInfo.Persons[0].ResidencePlace ?? string.Empty);

										docxDocument.Save();
									}
								}
							}
							catch (Exception exception)
							{
								excel.Quit();
								throw;
							}

							MessageBox.Show("Готово");
						}
					}
				}
			}
		}

		private static string MakeValidFileName(string name)
		{
			string invalidChars = System.Text.RegularExpressions.Regex.Escape(new string(System.IO.Path.GetInvalidFileNameChars()));
			string invalidRegStr = string.Format(@"([{0}]*\.+$)|([{0}]+)", invalidChars);

			return System.Text.RegularExpressions.Regex.Replace(name, invalidRegStr, "_");
		}
	}
}
