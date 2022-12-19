using Aspose.Words;
using Aspose.Words.Tables;

using System;
using System.Collections;
using System.Text.RegularExpressions;

using Paragraph = Aspose.Words.Paragraph;

namespace SubtitleEdit.Logic
{
	public static class Translation
	{
		public static string DocxToTxt(string file_name)
		{
			var document = new Document(file_name);
			DocumentBuilder builder = new DocumentBuilder(document);
			var newFileName = Regex.Replace(file_name, @"\.doc.$", ".tmp");

			Node[] tables = document.GetChildNodes(NodeType.Table, true).ToArray();
			foreach (Table table in tables)
			{
				Paragraph par = new Paragraph(document);

				table.ParentNode.InsertAfter(par, table);
				builder.MoveTo(par);
				builder.Font.Name = "Courier New";

				builder.Writeln(ConvertTable(table));
				table.Remove();
			}
			string documentText = document.Range.Text;
			return documentText
				.Replace("Evaluation Only. Created with Aspose.Words. Copyright 2003-2022 Aspose Pty Ltd.", "")
				.Replace("Created with an evaluation copy of Aspose.Words. To discover the full versions of our APIs please visit: https://products.aspose.com/words/", "");
		}

		private static string ConvertTable(Table tab)
		{
			string output = string.Empty;

			ArrayList columnWidhs = new ArrayList();
			int tableWidth = 0;
			string horizontalBorder = "";

			foreach (Row row in tab.Rows)
				foreach (Cell cell in row.Cells)
				{
					int cellIndex = row.Cells.IndexOf(cell);
					if (columnWidhs.Count > cellIndex)
					{
						if ((int) columnWidhs[cellIndex] < cell.GetText().Length)
							columnWidhs[cellIndex] = cell.GetText().Length;
					}
					else
						columnWidhs.Add(cell.GetText().Length);
				}

			//Calculate width of table
			for (int index = 0; index < columnWidhs.Count; index++)
				tableWidth += (int) columnWidhs[index];
			tableWidth += columnWidhs.Count;

			//for (int index = 0; index < tableWidth; index++)
			//	horizontalBorder += "-";
			//horizontalBorder += "\r\n";

			output += horizontalBorder;

			NodeCollection tableNotes = tab.GetChildNodes(NodeType.Paragraph, true);
			string regularTime = new Regex("^(0?[1-9]|1[0-2]):[0-5][0-9](:[0-5][0-9])?").ToString();
			string previousTimeCode = "00:00:00";
			foreach (Row row in tab.Rows)
			{
				string currentRow = "";

				foreach (Cell cell in row.Cells)
				{
					int cellIndex = row.Cells.IndexOf(cell);

					string curentCell = cell.GetText().Replace("\a", "").Replace("\n", "").Replace("\r", "");

					if (Regex.IsMatch(curentCell, regularTime))
					{
						var before = curentCell;
						curentCell = previousTimeCode + " --> " + curentCell;
						previousTimeCode = before;
					}

					if (cellIndex != row.Cells.Count - 1)
						curentCell += Environment.NewLine;
					currentRow += curentCell;
				}

				output += currentRow + Environment.NewLine;
			}

			return output;
		}
	}
}
