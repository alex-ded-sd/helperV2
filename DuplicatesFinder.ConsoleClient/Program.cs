namespace DuplicatesFinder.ConsoleClient
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.Linq;
	using Core;

	class Program
	{
		static void Main(string[] args) {
			ExcelWorker class1 = new ExcelWorker();
			List<Entity> allRecords = class1.ReadDataFrom(@"1.xlsx", "Січень");
            UniqueRecordsFinder finder = new UniqueRecordsFinder();
            List<Entity> uniqueRecords = finder.GetUniqueRecords(allRecords);
            //List<Entity> uniqueRecords = new List<Entity>();
            //while (result.Count != 0) {
            //	Entity item = result.FirstOrDefault();
            //	if (item != null) {
            //		Entity copyItem = new Entity(item);
            //		result.Remove(item);
            //		var duplicateRecords = result.FindAll(entity => entity.Equals(copyItem));
            //		foreach (Entity duplicateRecord in duplicateRecords) {
            //			copyItem.Count += duplicateRecord.Count;
            //			result.Remove(duplicateRecord);
            //		}
            //		uniqueRecords.Add(copyItem);
            //	}
            //}

            class1.WriteData(@"new.xlsx", uniqueRecords);
			Console.ReadLine();
		}
	}
}
