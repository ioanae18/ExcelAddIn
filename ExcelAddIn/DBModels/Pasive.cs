using System.ComponentModel.DataAnnotations.Schema;

namespace DBModels
{
	public class Pasive
	{
		public Pasive() { }

		public Pasive(string PasiveID, int CapProprii, int Provizioane, int DatLung, int DatScurt, int VenitAvans,
			string StoreID)
		{
			this.PasiveID = PasiveID;
			this.CapProprii = CapProprii;
			this.Provizioane = Provizioane;
			this.DatLung = DatLung;
			this.DatScurt = DatScurt;
			this.VenitAvans = VenitAvans;
			this.StoreID = StoreID;
		}

		public string PasiveID { get; set; }

		public int CapProprii { get; set; }

		public int Provizioane { get; set; }

		public int DatLung { get; set; }

		public int DatScurt { get; set; }

		public int VenitAvans { get; set; }

		public string StoreID { get; set; }
		[ForeignKey("StoreID")]
		public Store store { get; set; }
	}
}
