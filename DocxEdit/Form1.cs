using Nikse.SubtitleEdit.PluginLogic;

using SubtitleEdit.Logic;

using System.Linq;
using System.Windows.Forms;

namespace SubtitleEdit
{
	public partial class Form1: Form
	{
		public string PathToChosenFile { get; set; }
		private Subtitle Subtitle { get; set; } = new Subtitle();
		private SubRip SubRip { get; set; } = new SubRip();
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
				Filter = "Файл MS Word (*.docx; *.doc)|*.docx; *.doc"
			};
			if (openFileDialog.ShowDialog() == DialogResult.OK)
				PathToChosenFile = openFileDialog.FileName;
			else if (openFileDialog.ShowDialog() == DialogResult.Cancel)
				PathToChosenFile = "";
		}
	}
}
