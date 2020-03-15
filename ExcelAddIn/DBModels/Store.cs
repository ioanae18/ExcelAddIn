using System;

namespace DBModels
{
	public class Store
	{

		public Store() { }

		public Store(string StoreID, string StoreName, string Phone, string Email, string City, int Income)
		{
			this.StoreID = StoreID;
			this.StoreName = StoreName;
			this.Phone = Phone;
			this.Email = Email;
			this.City = City;
			this.Income = Income;
		}

		public string StoreID { get; set; }

		public string StoreName { get; set; }

		public string Phone { get; set; }

		public string Email { get; set; }

		public string City { get; set; }

		public int Income { get; set; }
	}
}
