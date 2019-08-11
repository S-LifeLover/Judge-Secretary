using System.IO;
using System.Windows;
using System.Windows.Forms;

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
				var result = fbd.ShowDialog();

				if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
				{
					string[] files = Directory.GetFiles(fbd.SelectedPath);

					System.Windows.Forms.MessageBox.Show("Files found: " + files.Length, "Message");
				}
			}
		}
	}
}
