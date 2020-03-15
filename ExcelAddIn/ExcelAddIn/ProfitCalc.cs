using DBModels;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ExcelAddIn
{
	public partial class ProfitCalc : Form
	{

		private const string _connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\Faculta\FACULTA_MASTER\ANULI-SEM1\Ingineria Programarii si Limbaje de Asamblare\ProiectFinalIP\ExcelAddIn\MagazinDB.accdb";

		private List<Store> listStore = new List<Store>();
		public ProfitCalc()
		{
			InitializeComponent();
		}

		private void BtnShow_Click(object sender, EventArgs e)
		{
			GetStore();

			int cost = int.Parse(txtBoxCost.Text);
			double profit = 0;
			foreach (var v in listStore)
			{
				profit = v.Income - cost;
			}

			Microsoft.Office.Interop.Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
			Microsoft.Office.Interop.Excel.Range range = Globals.ThisAddIn.Application.ActiveCell;

			int activeRow = range.Row;
			int activeColumn = range.Column;

			(sheet.Cells[activeRow, activeColumn] as Microsoft.Office.Interop.Excel.Range).Value2 = profit;
			this.Close();
		}

		public void GetStore()
		{
			using (OleDbConnection connection = new OleDbConnection(_connString))
			{
				connection.Open();
				OleDbDataReader reader = null;
				OleDbCommand command = new OleDbCommand("SELECT * FROM Store", connection);
				reader = command.ExecuteReader();
				while (reader.Read())
				{
					listStore.Add(new Store(reader[0].ToString(), reader[1].ToString(),
						reader[2].ToString(), reader[3].ToString(), reader[4].ToString(),
						int.Parse(reader[5].ToString()) ));
				}
			}
		}
	}
}
