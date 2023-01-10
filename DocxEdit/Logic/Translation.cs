using Spire.Doc;
using Spire.Doc.Collections;

using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Paragraph = Spire.Doc.Documents.Paragraph;

namespace SubtitleEdit.Logic
{
	public static class Translation
	{
		public static string DocxToString(string file_name)
		{
			Document document;
			try
			{
				document = new Document(file_name);
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

			TableCollection tables = document.Sections[0].Tables;
			foreach (Table table in tables)
				foreach (TableRow row in table.Rows)
					foreach (Paragraph paragraph in row.Cells[1].Paragraphs)
						paragraph.Text = "[" + paragraph.Text + "]";

			var newFileName = Regex.Replace(file_name, @"\.doc.?$", ".txt");
			document.SaveToFile(newFileName, FileFormat.Txt);

			return document.GetText();
		}

		public static string StringToDocx(string file_name, string subtitleText)
		{
			Document document;
			try
			{
				document = new Document(file_name);
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

			string[] lines = subtitleText
				.Split(Environment.NewLine.ToCharArray())
				.ToArray();
			for (int i = 1; i < lines.Length; i += 4)
				if (string.IsNullOrEmpty(lines[i]))
				{
					string nextLine = lines[i + 1];
					try
					{
						lines[i] = nextLine.Substring(1, nextLine.IndexOf(']') - 1);
					}
					catch
					{
						MessageBox.Show(
							$"Полностью или частично стрёрт маркер актёра. Реплика: {lines[i + 1]}.",
							"Ошибка сохранения файла.",
							MessageBoxButtons.OK,
							MessageBoxIcon.Error
						);
						return "";
					}
					lines[i + 1] = nextLine.Remove(0, nextLine.IndexOf(']') + 2);
				}
			lines = lines.Where(line => !string.IsNullOrEmpty(line)).ToArray();

			int j = 0;
			TableCollection tables = document.Sections[0].Tables;
			foreach (Table table in tables)
			{
				int startRowNumber = 0;
				if (table.Rows[0].Cells[0].Paragraphs[0].Text == "")
					startRowNumber = 1;

				for (int i = startRowNumber; i < table.Rows.Count; i++)
					foreach (TableCell cell in table.Rows[i].Cells)
					{
						// Т. к. в некоторых ячейках таблицы больше, чем два "параграфа" (переноса на другую строку),
						// а в массиве субтитров из SE символы "\n" просто заменяются на пробелы,
						// то для переноса текста из SE обратно в документ надо заменить все существующие параграфы одним.
						var paragraph = new Paragraph(document)
						{
							Text = lines[j++]
						};

						cell.Paragraphs.Clear();
						cell.Paragraphs.Add(paragraph);
					}
			}

			document.SaveToFile(file_name);
			return document.GetText();
		}
	}
}
