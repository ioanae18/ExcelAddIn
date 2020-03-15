using System;
using System.Collections.Generic;
using System.Text;

namespace DBModels
{
	public class Category
	{
		public Category() { }

		public Category(string CategoryID, string CategoryName)
		{
			this.CategoryID = CategoryID;
			this.CategoryName = CategoryName;
		}

		public string CategoryID { get; set; }
		public string CategoryName { get; set; }

	}
}
