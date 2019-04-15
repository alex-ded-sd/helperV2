namespace DuplicatesFinder.Core
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using OfficeOpenXml;

    public class ExcelWorker
	{
        public List<string> GetWorkSheets(string filePath)
        {
            List<string> sheetsName = new List<string>();
            FileInfo file = new FileInfo(filePath);
            using(ExcelPackage package = new ExcelPackage(file))
            {
                sheetsName.AddRange(package.Workbook.Worksheets.Select(sheet => sheet.Name));
            }

            return sheetsName;
        }

		public List<Entity> ReadDataFrom(string filePath, string sheetName) {
			List<Entity> records = new List<Entity>();
			FileInfo file = new FileInfo(filePath);
			using (ExcelPackage package = new ExcelPackage(file)) {
				ExcelWorksheet workSheet = package.Workbook.Worksheets[sheetName];
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
				ExcelWorksheet UniqueRecords = package.Workbook.Worksheets.Add("Уникальные треки");
				int row = 1;
				foreach (Entity uniqueRecord in uniqueRecords) {
					UniqueRecords.Cells[row, 1].Value = uniqueRecord.Singer;
					UniqueRecords.Cells[row, 2].Value = uniqueRecord.Name;
					UniqueRecords.Cells[row, 3].Value = uniqueRecord.Album;
					UniqueRecords.Cells[row, 4].Value = uniqueRecord.Count;
					row++;
				}
				package.Save(); 
			}
		}
	}
}