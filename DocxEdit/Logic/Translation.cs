
using Spire.Doc;
using Spire.Doc.Collections;
using Spire.Doc.Documents;

using System;
using System.Collections;
using System.Text.RegularExpressions;


namespace SubtitleEdit.Logic
{
	public static class Translation
	{
		public static string DocxToString(string file_name)
		{
			Document document = new Document(file_name);

			Section section = document.Sections[0];
			section.PageSetup.DifferentFirstPageHeaderFooter = true;
			section.HeadersFooters.FirstPageHeader.ChildObjects.Clear();
			section.HeadersFooters.Header.ChildObjects.Clear();

			var newFileName = Regex.Replace(file_name, @"\.doc.?$", ".txt");
			document.SaveToFile(newFileName, FileFormat.Txt);

			return document.GetText();
		}

		public static void StringToDocx(string file_name, string subtitleText)
		{
			var document = new Document(file_name);

			string[] lines = subtitleText.Split(Environment.NewLine.ToCharArray());
			TableCollection tables = document.Sections[0].Tables;

			int j = -1;
			foreach (Table table in tables)
				foreach (TableRow row in table.Rows)
					for (int i = 0; i < row.Cells.Count; ++i)
					{
						TableCell cell = row.Cells[i];

						foreach (Paragraph paragraph in cell.Paragraphs)
							paragraph.Text = lines[++j];
					}

			document.SaveToFile(file_name);
		}
	}
}
