using Microsoft.EntityFrameworkCore.Metadata.Internal;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;

namespace DBModels
{
	public class Product
	{
		private string productValue;

		public static void CalculateValue(Product product, Dictionary<string, string> pairs)
		{
			foreach (var key in pairs.Keys)
			{
				switch (key)
				{
					case nameof(productValue):
						product.productValue = pairs[key];
						break;
					default:
						break;
				}
			}
		}

		public Product() { }

		public Product(string ProductID, string ProductName, int Price, int ModelYear, string CategoryID)
		{
			this.ProductID = ProductID;
			this.ProductName = ProductName;
			this.Price = Price;
			this.ModelYear = ModelYear;
			this.CategoryID = CategoryID;
		}

		public string ProductID { get; set; }

		public string ProductName { get; set; }

		public int Price { get; set; }

		public int ModelYear { get; set; }

		public string CategoryID { get; set; }

		[ForeignKey("CategoryID")]
		public Category category { get; set; }

		public override string ToString()
		{
			return $"{productValue} {Price} {ProductID} {ProductName} {ModelYear}";
		}
	}
}
