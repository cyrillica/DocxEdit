﻿using Spire.Doc;
using Spire.Doc.Collections;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Paragraph = Spire.Doc.Documents.Paragraph;

namespace DocxEdit.Logic
{
	public static class Translation
	{
		public static string DocxToString(string fileName)
		{
			Document document;
			try
			{
				document = new Document(fileName);
			}
			catch
			{
				MessageBox.Show(
					"Файл занят сторонним процессом. Завершите его, прежде чем продолжить работу.",
					"Невозможно открыть файл.",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error
				);
				return "";
			}

			Section section = document.Sections[0];
			section.PageSetup.DifferentFirstPageHeaderFooter = true;
			section.HeadersFooters.FirstPageHeader.ChildObjects.Clear();
			section.HeadersFooters.Header.ChildObjects.Clear();

			string documentText = "";
			TableCollection tables = document.Sections[0].Tables;
			foreach (Table table in tables)
				foreach (TableRow row in table.Rows)
					if (row.Cells[0].Paragraphs[0].Text == "")
					{
						table.Rows.Remove(row);
						break;
					}

			foreach (Table table in tables)
				foreach (TableRow row in table.Rows)
					foreach (TableCell cell in row.Cells)
					{
						cell.LastParagraph.Text += '|';
						for (int i = 0; i < cell.Paragraphs.Count; ++i)
							documentText += cell.Paragraphs[i].Text;
					}

			string newFileName = Regex.Replace(fileName, @"\.doc.?$", ".ass");
			File.WriteAllText(newFileName, RawTextToASSA(documentText));
			//document.SaveToFile(newFileName, FileFormat.Txt);

			return File.ReadAllText(newFileName);
		}

		static string RawTextToASSA(string rawText)
		{
			const int optimalCharacterRatePerSecond = 15;
			string result = @"[Script Info]
; This is an Advanced Sub Station Alpha v4+ script.
Title: Untitled
ScriptType: v4.00+

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
Style: Default,Arial,20,&H00FFFFFF,&H0000FFFF,&H00000000,&H00000000,0,0,0,0,100,100,0,0,1,1,1,2,10,10,10,1

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text";
			result += Environment.NewLine;

			List<string> rawLines = rawText
				.Split(('|' + Environment.NewLine).ToCharArray())
				.Where(line => !string.IsNullOrEmpty(line))
				.ToList();

			for (int i = 0; i < rawLines.Count - 2; i += 3)
			{
				TimeOnly startTime = new TimeOnly(
					0,
					int.Parse(rawLines[i].Substring(0, 2)),
					int.Parse(rawLines[i].Substring(3, 2)),
					0,
					0
				);
				int lettersCount = Regex.Matches(rawLines[i + 2], @"\w").Count; // кол-во буквенных символов для расчёта времени конца реплики
				TimeOnly stopTime = startTime.Add(new TimeSpan(0, 0, lettersCount / optimalCharacterRatePerSecond));

				string line = $"Dialogue: 0,{startTime:H:mm:ss.ff},{stopTime:H:mm:ss.ff},Default,{rawLines[i + 1]},0,0,0,,{rawLines[i + 2]}" + Environment.NewLine;
				result += line;
			}

			return result;
		}

		public static string ASSAToDocx(string fileName, string subtitleText)
		{
			Document document;
			try
			{
				document = new Document(fileName);
			}
			catch
			{
				MessageBox.Show(
					"Файл занят сторонним процессом. Завершите его, прежде чем продолжить работу.",
					"Невозможно открыть файл.",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error
				);
				return "";
			}

			List<string> dialogues = subtitleText
				.Split(Environment.NewLine.ToCharArray())
				.Where(line => !string.IsNullOrEmpty(line) && line.Contains("Dialogue"))
				.ToList();
			List<List<string>> contentForTable = new List<List<string>>();

			foreach (string dialogue in dialogues)
			{
				int indexOfText = FindNthOccur(dialogue, ',', 9) + 1;
				string text = dialogue.Substring(indexOfText);
				string dialogueWithoutText = dialogue.Remove(indexOfText - 1);

				string start = $"{dialogueWithoutText.Split(',')[1]:mm:ss}";
				string name = dialogueWithoutText.Split(',')[4];

				contentForTable.Add(new List<string>() { start, name, text });
			}

			TableCollection tables = document.Sections[0].Tables;
			foreach (Table table in tables)
			{
				int startRowNumber = 0;
				if (table.Rows[0].Cells[0].Paragraphs[0].Text == "")
					startRowNumber = 1;

				for (int i = startRowNumber; i < table.Rows.Count; i++)
				{
					int j = 0;
					foreach (TableCell cell in table.Rows[i].Cells)
					{
						/* Т. к. в некоторых ячейках таблицы больше, чем два "параграфа" (переноса на другую строку),
						* а в массиве субтитров из SE символы "\n" просто заменяются на пробелы,
						* то для переноса текста из SE обратно в документ надо заменить все существующие параграфы одним.
						*/
						var paragraph = new Paragraph(document)
						{
							Text = contentForTable[startRowNumber == 0 ? i : i - 1][j++]
						};

						cell.Paragraphs.Clear();
						cell.Paragraphs.Add(paragraph);
					}
				}
			}

			document.SaveToFile(fileName);
			return subtitleText;
		}

		static int FindNthOccur(string str, char ch, int n)
		{
			int occur = 0;

			for (int i = 0; i < str.Length; i++)
			{
				if (str[i] == ch)
					occur++;
				if (occur == n)
					return i;
			}
			return -1;
		}
	}
}
