using SubtitleEdit;
using SubtitleEdit.Logic;

using System;
using System.Collections.Generic;
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
			OpenFileDialog openFileDialog = new OpenFileDialog
			{
				Filter = "Файлы MS Word (*.docx; *.doc)|*.docx; *.doc"
			};
			Translation.DocxToTxt(openFileDialog.FileName);
			//subtitle = subtitle.Trim();
			//if (string.IsNullOrEmpty(subtitle))
			//{
			//	MessageBox.Show("Субтитры не загружены", parentForm.Text,
			//		MessageBoxButtons.OK, MessageBoxIcon.Warning);
			//	return string.Empty;
			//}

			if (!string.IsNullOrEmpty(listViewLineSeparatorString))
				Configuration.ListViewLineSeparatorString = listViewLineSeparatorString;

			var list = new List<string>();
			foreach (string line in subtitle.Replace(Environment.NewLine, "\n").Split('\n'))
				list.Add(line);

			var sub = new Subtitle();
			var srt = new SubRip();
			srt.LoadSubtitle(sub, list, subtitleFileName);
			using (var form = new MainForm(sub, (this as IPlugin).Text, (this as IPlugin).Description, parentForm))
			{
				if (form.ShowDialog(parentForm) == DialogResult.OK)
					return form.FixedSubtitle;
			}
			return string.Empty;
		}
	}
}