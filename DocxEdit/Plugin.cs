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
			if (form.PathToChosenFile == null)
				return "";

			subtitle = Translation.DocxToTxt(form.PathToChosenFile).Trim();
			subtitleFileName = form.PathToChosenFile;
			var sub = new Subtitle();
			var srt = new SubRip();
			srt.LoadSubtitle(sub, subtitle.SplitToLines().ToList(), form.PathToChosenFile);

			Configuration.CurrentFrameRate = 70;

			if (!string.IsNullOrEmpty(listViewLineSeparatorString))
				Configuration.ListViewLineSeparatorString = listViewLineSeparatorString;

			return subtitle;
			//return srt.ToText(sub, subtitleFileName);
			//return "1\r00:00:00,498 --> 00:00:02,827\r- Here's what I love most\rabout food and diet.\r\r2\r00:00:02,827 --> 00:00:06,383\rWe all eat several times a day,\rand we're totally in charge\r\r3\r00:00:06,383 --> 00:00:09,427\rof what goes on our plate\rand what stays off.";
		}
	}
}