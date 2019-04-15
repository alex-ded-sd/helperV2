namespace DuplicatesFinder.Core
{
	using System;
	using System.Collections.Generic;
	using System.Diagnostics;
	using System.IO;
	using OfficeOpenXml;

	public class Entity
	{
		public Entity() {
			
		}
		public Entity(Entity item) {
			Singer = item.Singer.Clone().ToString();
			Name = item.Name.Clone().ToString();
			Album = item.Album.Clone().ToString();
			Count = item.Count;
		}

		public override bool Equals(object obj) {
			return Equals((Entity)obj);
		}

		protected bool Equals(Entity other) {
			return string.Equals(Singer, other.Singer, StringComparison.OrdinalIgnoreCase) && string.Equals(Name, other.Name, StringComparison.OrdinalIgnoreCase) && string.Equals(Album, other.Album, StringComparison.OrdinalIgnoreCase);
		}


		public static bool operator ==(Entity left, Entity right) {
			return Equals(left, right);
		}

		public static bool operator !=(Entity left, Entity right) {
			return !Equals(left, right);
		}

		public string Singer { get; set; }

		public string Name { get; set; }

		public string Album { get; set; }

		public double Count { get; set; }

	}

	public class Class1
	{
		public List<Entity> ReadDataFrom(string filePath) {
			List<Entity> records = new List<Entity>();
			FileInfo file = new FileInfo(filePath);
			using (ExcelPackage package = new ExcelPackage(file)) {
				ExcelWorksheet workSheet = package.Workbook.Worksheets["Січень"];
				int totalRows = workSheet.Dimension.Rows;
				for (int i = 2; i <= totalRows; i++) {
					records.Add(new Entity {
						Singer = workSheet.Cells[i, 2].Value.ToString(),
						Name = workSheet.Cells[i, 3].Value.ToString(),
						Album = workSheet.Cells[i, 4].Value.ToString(),
						Count = (double)workSheet.Cells[i, 6].Value,
					});
				}
			}

			return records;
		}

		public void WriteData(string filePath, List<Entity> uniqueRecords) {
			List<Entity> records = new List<Entity>();
			FileInfo file = new FileInfo(filePath);
			using (ExcelPackage package = new ExcelPackage(file)) {
				ExcelWorksheet workSheet = package.Workbook.Worksheets["Січень"];
				string header1 = workSheet.Cells[1, 1].Value.ToString();
				string header2 = workSheet.Cells[1, 2].Value.ToString();
				string header3 = workSheet.Cells[1, 3].Value.ToString();
				string header4 = workSheet.Cells[1, 4].Value.ToString();
				string header5 = workSheet.Cells[1, 5].Value.ToString();
				string header6 = workSheet.Cells[1, 6].Value.ToString();
				ExcelWorksheet UniqueRecords = package.Workbook.Worksheets.Add("Уникальные треки");
				UniqueRecords.Cells[1, 1].Value = header1;
				UniqueRecords.Cells[1, 2].Value = header2;
				UniqueRecords.Cells[1, 3].Value = header3;
				UniqueRecords.Cells[1, 4].Value = header4;
				UniqueRecords.Cells[1, 5].Value = header5;
				UniqueRecords.Cells[1, 6].Value = header6;
				int row = 2;
				foreach (Entity uniqueRecord in uniqueRecords) {
					UniqueRecords.Cells[row, 2].Value = uniqueRecord.Singer;
					UniqueRecords.Cells[row, 3].Value = uniqueRecord.Name;
					UniqueRecords.Cells[row, 4].Value = uniqueRecord.Album;
					UniqueRecords.Cells[row, 6].Value = uniqueRecord.Count;
					row++;
				}
				package.Save(); 
			}
		}
	}
}