using SubtitleEdit;
using SubtitleEdit.Logic;

using System.Linq;
using System.Windows.Forms;

namespace Nikse.SubtitleEdit.PluginLogic
{
	public class DocxEdit : IPlugin // dll file name must "<classname>.dll" - e.g. "Haxor.dll"
	{
		string IPlugin.Name => "DocxEdit";

		string IPlugin.Text => "DocxEdit (открыть файл)";

		decimal IPlugin.Version => 0.1M;

		string IPlugin.Description => "Edit MS Word files";

		// Can be one of these: file, tool, sync, translate, spellcheck
		string IPlugin.ActionType => "file";

		string IPlugin.Shortcut => string.Empty;

		string IPlugin.DoAction(Form parentForm, string subtitle, double frameRate, string listViewLineSeparatorString, string subtitleFileName, string videoFileName, string rawText)
		{
			var form = new Form1((this as IPlugin).Name, (this as IPlugin).Description);

			subtitle = Translation.DocxToTxt(form.PathToChosenFile).Trim();
			subtitleFileName = form.PathToChosenFile;
			var sub = new Subtitle();
			var srt = new SubRip();
			srt.LoadSubtitle(sub, subtitle.SplitToLines().ToList(), form.PathToChosenFile);

			Configuration.CurrentFrameRate = frameRate;

			if (!string.IsNullOrEmpty(listViewLineSeparatorString))
				Configuration.ListViewLineSeparatorString = listViewLineSeparatorString;

			return subtitle;
		}
	}
}