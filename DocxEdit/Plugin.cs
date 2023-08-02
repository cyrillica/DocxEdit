﻿using DocxEdit.Logic;

using SubtitleEdit;

using System.Windows.Forms;

namespace Nikse.SubtitleEdit.PluginLogic
{
	public class DocxEdit: IPlugin // dll file name must "<classname>.dll" - e.g. "Haxor.dll"
	{
		string IPlugin.Name => "DocxEdit";
		string IPlugin.Text => "DocxEdit (открыть или сохранить файл)";
		decimal IPlugin.Version => 0.1M;
		string IPlugin.Description => "Edit Office files";
		// Can be one of these: file, tool, sync, translate, spellcheck
		string IPlugin.ActionType => "file";
		string IPlugin.Shortcut => string.Empty;

		string IPlugin.DoAction(Form parentForm, string subtitle, double frameRate, string listViewLineSeparatorString, string subtitleFileName, string videoFileName, string rawText)
		{
			Configuration.CurrentFrameRate = frameRate;
			if (!string.IsNullOrEmpty(listViewLineSeparatorString))
				Configuration.ListViewLineSeparatorString = listViewLineSeparatorString;

			string outputRawText;

			if (subtitle == "\r\n\r\n")
			{
				var form = new Form1((this as IPlugin).Name, (this as IPlugin).Description);
				if (form.PathToChosenFile == null)
					return "";
				subtitleFileName = form.PathToChosenFile;
				parentForm.Text = subtitleFileName;

				outputRawText = Translation.DocxToPlainText(form.PathToChosenFile);
			}
			else
			{
				outputRawText = Translation.SRTToDocx(parentForm.Text.Replace("*", ""), rawText);
				if (subtitle != "")
					parentForm.Text = parentForm.Text.Replace("*", "");
			}

			return outputRawText;
		}
	}
}