namespace DuplicatesFinder.Core
{
    using System;

    public class Entity
	{
		public Entity() {
			
		}
		public Entity(Entity item) {
			Singer = item.Singer.Clone().ToString();
			Name = item.Name.Clone().ToString();
			Album = item.Album.Clone().ToString();
			IsrcCode = item.IsrcCode.Clone().ToString();
			Count = item.Count;
		}

		public override bool Equals(object obj) {
			return Equals((Entity)obj);
		}

		protected bool Equals(Entity other) {
			return string.Equals(Singer, other.Singer, StringComparison.OrdinalIgnoreCase)
			       && string.Equals(Name, other.Name, StringComparison.OrdinalIgnoreCase)
			       && string.Equals(IsrcCode, other.IsrcCode, StringComparison.OrdinalIgnoreCase)
			       && string.Equals(Album, other.Album, StringComparison.OrdinalIgnoreCase);
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

		public string IsrcCode { get; set; }

		public double Count { get; set; }

	}
}