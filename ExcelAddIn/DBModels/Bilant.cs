using System.ComponentModel.DataAnnotations.Schema;

namespace DBModels
{
	public class Bilant
	{
		public Bilant() { }

		public Bilant(int BilantID, string ActiveID, int ImobNecorp, int ImobCorp, int ImobFin, int Stocuri, int Creante,
			int CasaConturi, string PasiveID, int CapProprii, int Provizioane, int DatLung, int DatScurt, int VenitAvans,
			string StoreID)
		{
			this.BilantID = BilantID;
			this.ActiveID = ActiveID;
			this.ImobNecorp = ImobNecorp;
			this.ImobCorp = ImobCorp;
			this.ImobFin = ImobFin;
			this.Stocuri = Stocuri;
			this.Creante = Creante;
			this.CasaConturi = CasaConturi;
			this.PasiveID = PasiveID;
			this.CapProprii = CapProprii;
			this.Provizioane = Provizioane;
			this.DatLung = DatLung;
			this.DatScurt = DatScurt;
			this.VenitAvans = VenitAvans;
			this.StoreID = StoreID;
		}

		public int BilantID { get; set; }
		public string ActiveID { get; set; }
		[ForeignKey("ActiveID")]
		public Active active { get; set; }
		public int ImobNecorp { get; set; }
		public int ImobCorp { get; set; }
		public int ImobFin { get; set; }
		public int Stocuri { get; set; }
		public int Creante { get; set; }
		public int CasaConturi { get; set; }
		public string PasiveID { get; set; }
		[ForeignKey("PasiveID")]
		public Pasive passive { get; set; }
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
