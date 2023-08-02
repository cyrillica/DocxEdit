using System.Windows.Forms;

namespace SubtitleEdit
{
	public partial class Form1: Form
	{
		public string PathToChosenFile { get; set; }
		public string StringSubtitle { get; set; }

		public Form1()
		{
			InitializeComponent();
		}

		public Form1(string name, string description)
		{
			InitializeComponent();
			Text = name;

			OpenFileDialog openFileDialog = new OpenFileDialog
			{
				Filter = "Файл Office (*.docx; *.doc; *.odt)|*.docx; *.doc; *.odt"
			};
			if (openFileDialog.ShowDialog() == DialogResult.OK)
				PathToChosenFile = openFileDialog.FileName;
		}
	}
}
