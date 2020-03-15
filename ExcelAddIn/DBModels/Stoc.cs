using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel.DataAnnotations.Schema;
using Microsoft.EntityFrameworkCore.Metadata.Internal;

namespace DBModels
{
	class Stoc
	{
		public Stoc() { }

		public Stoc(string StoreID, string ProductID, int Quantity)
		{
			this.StoreID = StoreID;
			this.ProductID = ProductID;
			this.Quantity = Quantity;
		}

		public string StoreID { get; set; }

		[ForeignKey("StoreID")]
		public Stoc store { get; set; }
	
		public string ProductID { get; set; }

		[ForeignKey("ProductID")]
		public Product product { get; set; }

		public int Quantity { get; set; }
	}
}
