namespace DuplicatesFinder.Core
{
	using System;
	using System.Collections.Generic;
	using System.Diagnostics;
	using System.IO;
    using System.Linq;
    using OfficeOpenXml;

    public class ExcelWorker
    {
	    private readonly string _filePath;
	    private readonly string _errorMessage = "Екселька почему-то не загрузилась. Загрузи кнопкой и попробуй ещё раз";

		public ExcelWorker(string filePath) {
			_filePath = filePath;
		}

        public List<string> GetWorkSheets()
        {
	        List<string> sheetsName = new List<string>();
			FileInfo file = new FileInfo(_filePath);
	        using (ExcelPackage package = new ExcelPackage(file)) {
		        sheetsName.AddRange(package.Workbook.Worksheets.Select(sheet => sheet.Name));
	        }
	        return sheetsName;
        }


		public List<Entity> ReadDataFrom(string sheetName) {
			FileInfo file = new FileInfo(_filePath);
			List<Entity> records = new List<Entity>();
			using (ExcelPackage package = new ExcelPackage(file)) {
				ExcelWorksheet workSheet = package.Workbook.Worksheets[sheetName];
				int totalRows = workSheet.Dimension.Rows;
				for (int i = 2; i <= totalRows; i++) {
					records.Add(new Entity
					{
						Singer = workSheet.Cells[i, 2].Value.ToString(),
						Name = workSheet.Cells[i, 3].Value.ToString(),
						Album = workSheet.Cells[i, 4].Value.ToString(),
						Count = (double) workSheet.Cells[i, 6].Value,
					});
				}
			}
			return records;
		}

		public void SaveAndShowAsExcelFile(List<Entity> uniqueRecords) {
			string fileName = "обработан.xlsx";
			using (ExcelPackage package = new ExcelPackage()) {
				ExcelWorksheet uniqueRecordsSheet = package.Workbook.Worksheets.Add("Уникальные треки");
				int row = 1;
				foreach (Entity uniqueRecord in uniqueRecords) {
					uniqueRecordsSheet.Cells[row, 1].Value = uniqueRecord.Singer;
					uniqueRecordsSheet.Cells[row, 2].Value = uniqueRecord.Name;
					uniqueRecordsSheet.Cells[row, 3].Value = uniqueRecord.Album;
					uniqueRecordsSheet.Cells[row, 4].Value = uniqueRecord.Count;
					row++;
				}
				
				package.SaveAs(new FileInfo(fileName));
			}
			Process process = new Process {StartInfo = {FileName = fileName}};
			process.Start();
		}
	}
}