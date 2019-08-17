using System.IO;
using System.Windows;
using System.Windows.Forms;
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

					fbd.Description = "Выберете папку, куда сохранить постановления";
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
						}
					}

					MessageBox.Show("Готово");
				}
			}
		}
	}
}
