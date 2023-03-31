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
		public static string DocxToPlainText(string fileName)
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

			Regex mmss = new Regex(@"(([0]?[0-9][0-9]|[0-9]):([0-5][0-9]))"); //время типа "МинМин:СекСек"
			TableCollection tables = document.Sections[0].Tables;
			// вставка субтитра с таймкодом в реплике
			foreach (Table table in tables)
				for (int i = 0; i < table.Rows.Count; ++i)
				{
					string cellText = "";
					foreach (Paragraph paragraph in table.Rows[i].Cells[2].Paragraphs)
					{
						int t = table.Rows[i].Cells.Count;
						paragraph.Text = Regex.Replace(paragraph.Text, @"\^|\/", "");
						cellText += paragraph.Text + " ";
					}

					List<string> timecodes = new List<string>();
					foreach (Match match in mmss.Matches(cellText))
						timecodes.Add(match.Value);

					if (timecodes.Count == 0)
						continue;

					string[] replics = cellText.Split(timecodes.ToArray(), StringSplitOptions.None);

					table.Rows[i].Cells[2].Paragraphs.Clear();
					table.Rows[i].Cells[2].AddParagraph().Text = replics[0];

					string author = table.Rows[i].Cells[1].FirstParagraph.Text;

					int j = 1;
					foreach (string timecode in timecodes)
					{
						TableRow trow = table.AddRow(false);

						trow.Cells[0].AddParagraph().Text = timecode;
						trow.Cells[1].AddParagraph().Text = author;
						trow.Cells[2].AddParagraph().Text = replics[j++];

						table.Rows.Insert(++i, trow);
						document.SaveToFile(fileName);
					}
				}

			Section section = document.Sections[0];
			section.PageSetup.DifferentFirstPageHeaderFooter = true;
			section.HeadersFooters.FirstPageHeader.ChildObjects.Clear();
			section.HeadersFooters.Header.ChildObjects.Clear();

			string documentText = "";
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

			string newFileName = Regex.Replace(fileName, @"\.doc.?$", ".txt");
			File.WriteAllText(newFileName, PlainTextToSRT(documentText));

			return File.ReadAllText(newFileName);
		}

		static string PlainTextToSRT(string rawText)
		{
			const int optimalCharacterRatePerSecond = 15;
			string result = "";

			List<string> rawLines = rawText
				.Split(('|' + Environment.NewLine).ToCharArray())
				.Where(line => !string.IsNullOrEmpty(line))
				.ToList();

			for (int i = 0, j = 1; i < rawLines.Count - 2; i += 3, ++j)
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

				string line = $"{j}\n{startTime:HH:mm:ss,fff} --> {stopTime:HH:mm:ss,fff}\n{rawLines[i + 2]}\n\n";
				result += line;
			}

			return result;
		}

		public static string SRTToDocx(string fileName, string subtitleText)
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

			List<string> dialogues = Regex
				.Split(subtitleText, @"\d+\r\n") // разделение по цифрам
				.Where(line => !string.IsNullOrEmpty(line))
				.ToList();

			for (int i = 0; i < dialogues.Count; i += 2)
			{
				TimeOnly startTime = new TimeOnly(
					0,
					int.Parse(dialogues[i].Substring(3, 2)),
					int.Parse(dialogues[i].Substring(6, 2)),
					0,
					0
				);
				dialogues[i] = $"{startTime:mm:ss}";
			}

			TableCollection tables = document.Sections[0].Tables;
			foreach (Table table in tables)
			{
				int startRowNumber = 0;
				if (table.Rows[0].Cells[0].Paragraphs[0].Text == "")
					startRowNumber = 1;

				int j = 0;
				for (int i = startRowNumber; i < table.Rows.Count; ++i)
				{
					/* Т. к. в некоторых ячейках таблицы больше, чем два "параграфа" (переноса на другую строку),
					* а в массиве субтитров из SE символы "\n" просто заменяются на пробелы,
					* то для переноса текста из SE обратно в документ надо заменить все существующие параграфы одним.
					*/
					var startTimeParagraph = new Paragraph(document)
					{
							Text = dialogues[j++]
					};

					table.Rows[i].Cells[0].Paragraphs.Clear();
					table.Rows[i].Cells[0].Paragraphs.Add(startTimeParagraph);

					var replicTimeParagraph = new Paragraph(document)
					{
						Text = dialogues[j++]
					};

					table.Rows[i].Cells[2].Paragraphs.Clear();
					table.Rows[i].Cells[2].Paragraphs.Add(replicTimeParagraph);
				}
			}

			string srtFileName = Regex.Replace(fileName, @"\.doc.?$", ".srt");
			File.WriteAllText(srtFileName, subtitleText);
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
