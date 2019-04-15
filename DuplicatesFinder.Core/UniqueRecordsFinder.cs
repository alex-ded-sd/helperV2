using System.Collections.Generic;
using System.Linq;

namespace DuplicatesFinder.Core
{
    public class UniqueRecordsFinder
    {
        public List<Entity> GetUniqueRecords(List<Entity> allRecords)
        {
            List<Entity> uniqueRecords = new List<Entity>();
            while(allRecords.Count != 0)
            {
                Entity item = allRecords.FirstOrDefault();
                if(item != null)
                {
                    Entity copyItem = new Entity(item);
                    allRecords.Remove(item);
                    var duplicateRecords = allRecords.FindAll(entity => entity.Equals(copyItem));
                    foreach(Entity duplicateRecord in duplicateRecords)
                    {
                        copyItem.Count += duplicateRecord.Count;
                        allRecords.Remove(duplicateRecord);
                    }
                    uniqueRecords.Add(copyItem);
                }
            }

            return uniqueRecords;
        }
    }
}