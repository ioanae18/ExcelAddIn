using System.ComponentModel.DataAnnotations.Schema;

namespace DBModels
{
	public class Active
	{
		public Active() { }

		public Active(string ActiveID, int ImobNecorp, int ImobCorp, int ImobFin, int Stocuri, int Creante, int CasaConturi,
			string StoreID)
		{
			this.ActiveID = ActiveID;
			this.ImobNecorp = ImobNecorp;
			this.ImobCorp = ImobCorp;
			this.ImobFin = ImobFin;
			this.Stocuri = Stocuri;
			this.Creante = Creante;
			this.CasaConturi = CasaConturi;
			this.StoreID = StoreID;
		}

		public string ActiveID { get; set; }
		public int ImobNecorp { get; set; }
		public int ImobCorp { get; set; }
		public int ImobFin { get; set; }
		public int Stocuri { get; set; }
		public int Creante { get; set; }
		public int CasaConturi { get; set; }
		public string StoreID { get; set; }
		[ForeignKey("StoreID")]
		public Store store { get; set; }
	}
}
