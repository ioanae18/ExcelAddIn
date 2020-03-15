using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using DBModels;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn
{
	public partial class Ribbon1
	{

		private const string _connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\Faculta\FACULTA_MASTER\ANULI-SEM1\Ingineria Programarii si Limbaje de Asamblare\ProiectFinalIP\ExcelAddIn\MagazinDB.accdb";

		private List<Active> listActive = new List<Active>();
		private List<Pasive> listPasive = new List<Pasive>();
		private List<Bilant> listBilant = new List<Bilant>();
		private List<Store> listStore = new List<Store>();
		private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
		{

		}

		private void Button1_Click(object sender, RibbonControlEventArgs e)
		{

			GetDataFromDataBase();
			AddInExcelDatas();
			listActive = new List<Active>();
			listPasive = new List<Pasive>();
			listBilant = new List<Bilant>();
			listStore = new List<Store>();
		}

		private void AddInExcelDatas()
		{
			Microsoft.Office.Interop.Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
			Microsoft.Office.Interop.Excel.Range range = Globals.ThisAddIn.Application.ActiveCell;

			int activeRow = range.Row;
			int activeColumn = range.Column;

			AddHeader();
			int i = 0;
			foreach (var b in listBilant)
			{
				(sheet.Cells[activeRow + i, activeColumn] as Microsoft.Office.Interop.Excel.Range).Value2 = b.BilantID;
				(sheet.Cells[activeRow + i, activeColumn + 1] as Microsoft.Office.Interop.Excel.Range).Value2 = b.ImobNecorp;
				(sheet.Cells[activeRow + i, activeColumn + 2] as Microsoft.Office.Interop.Excel.Range).Value2 = b.ImobCorp;
				(sheet.Cells[activeRow + i, activeColumn + 3] as Microsoft.Office.Interop.Excel.Range).Value2 = b.ImobFin;
				(sheet.Cells[activeRow + i, activeColumn + 4] as Microsoft.Office.Interop.Excel.Range).Value2 = b.Stocuri;
				(sheet.Cells[activeRow + i, activeColumn + 5] as Microsoft.Office.Interop.Excel.Range).Value2 = b.Creante;
				(sheet.Cells[activeRow + i, activeColumn + 6] as Microsoft.Office.Interop.Excel.Range).Value2 = b.CasaConturi;
				(sheet.Cells[activeRow + i, activeColumn + 7] as Microsoft.Office.Interop.Excel.Range).Value2 = b.CapProprii;
				(sheet.Cells[activeRow + i, activeColumn + 8] as Microsoft.Office.Interop.Excel.Range).Value2 = b.Provizioane;
				(sheet.Cells[activeRow + i, activeColumn + 9] as Microsoft.Office.Interop.Excel.Range).Value2 = b.DatLung;
				(sheet.Cells[activeRow + i, activeColumn + 10] as Microsoft.Office.Interop.Excel.Range).Value2 = b.DatScurt;
				(sheet.Cells[activeRow + i, activeColumn + 11] as Microsoft.Office.Interop.Excel.Range).Value2 = b.VenitAvans;
				(sheet.Cells[activeRow + i, activeColumn + 12] as Microsoft.Office.Interop.Excel.Range).Value2 = b.StoreID;
				i++;
				sheet.Columns.AutoFit();
			}
		}

		private void AddHeader()
		{
			Microsoft.Office.Interop.Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
			Microsoft.Office.Interop.Excel.Range range = Globals.ThisAddIn.Application.ActiveCell;

			int activeRow = range.Row;
			int activeColumn = range.Column;

			(sheet.Cells[activeRow, activeColumn] as Microsoft.Office.Interop.Excel.Range).Value2 = "Bilant ID";
			(sheet.Cells[activeRow, activeColumn + 1] as Microsoft.Office.Interop.Excel.Range).Value2 = "Imobilizari necorporale";
			(sheet.Cells[activeRow, activeColumn + 2] as Microsoft.Office.Interop.Excel.Range).Value2 = "Imobilizari corporale";
			(sheet.Cells[activeRow, activeColumn + 3] as Microsoft.Office.Interop.Excel.Range).Value2 = "Imobilizari financiare";
			(sheet.Cells[activeRow, activeColumn + 4] as Microsoft.Office.Interop.Excel.Range).Value2 = "Stocuri";
			(sheet.Cells[activeRow, activeColumn + 5] as Microsoft.Office.Interop.Excel.Range).Value2 = "Creante";
			(sheet.Cells[activeRow, activeColumn + 6] as Microsoft.Office.Interop.Excel.Range).Value2 = "Casa si conturi la banci";
			(sheet.Cells[activeRow, activeColumn + 7] as Microsoft.Office.Interop.Excel.Range).Value2 = "Capitaluri proprii";
			(sheet.Cells[activeRow, activeColumn + 8] as Microsoft.Office.Interop.Excel.Range).Value2 = "Provizioane";
			(sheet.Cells[activeRow, activeColumn + 9] as Microsoft.Office.Interop.Excel.Range).Value2 = "Datorii pe termen lung";
			(sheet.Cells[activeRow, activeColumn + 10] as Microsoft.Office.Interop.Excel.Range).Value2 = "Datorii pe termen scurt";
			(sheet.Cells[activeRow, activeColumn + 11] as Microsoft.Office.Interop.Excel.Range).Value2 = "Venituri in avans";
			(sheet.Cells[activeRow, activeColumn + 12] as Microsoft.Office.Interop.Excel.Range).Value2 = "Firma";
			sheet.Columns.AutoFit();
		}

		public void GetDataFromDataBase()
		{
			GetActive();
			GetPasive();
			GetBilant();
			GetStore();
			foreach (var a in listActive)
			{
				a.store = listStore.Where(x => x.StoreID == a.ActiveID).FirstOrDefault();
			}

			foreach (var b in listBilant)
			{
				b.active = listActive.Where(x => x.ActiveID == b.BilantID.ToString()).FirstOrDefault();
				b.passive = listPasive.Where(x => x.PasiveID == b.BilantID.ToString()).FirstOrDefault();
			}
		}

		public void GetBilantBy(int id)
		{
			using (OleDbConnection connection = new OleDbConnection(_connString))
			{
				connection.Open();
				OleDbDataReader reader = null;
				OleDbCommand command = new OleDbCommand("SELECT * FROM Bilant WHERE BilantID =  @id", connection);
				command.Parameters.AddWithValue("@BilantID", id);
				reader = command.ExecuteReader();
				while (reader.Read())
				{
					listBilant.Add(new Bilant(int.Parse(reader[0].ToString()), reader[1].ToString(), 
						int.Parse(reader[2].ToString()), int.Parse(reader[3].ToString()), int.Parse(reader[4].ToString()),
						int.Parse(reader[5].ToString()), int.Parse(reader[6].ToString()), int.Parse(reader[7].ToString()), 
						reader[8].ToString(), int.Parse(reader[9].ToString()), int.Parse(reader[9].ToString()), int.Parse(reader[9].ToString()),
						 int.Parse(reader[9].ToString()), int.Parse(reader[9].ToString()), reader[9].ToString() ));
				}
			}
		}

		public void GetActive()
		{
			using (OleDbConnection connection = new OleDbConnection(_connString))
			{
				connection.Open();
				OleDbDataReader reader = null;
				OleDbCommand command = new OleDbCommand("SELECT * FROM Active", connection);
				reader = command.ExecuteReader();
				while (reader.Read())
				{
					listActive.Add(new Active(reader[0].ToString(), int.Parse(reader[1].ToString()),
						int.Parse(reader[2].ToString()), int.Parse(reader[3].ToString()), int.Parse(reader[4].ToString()),
						int.Parse(reader[5].ToString()), int.Parse(reader[6].ToString()), reader[7].ToString()));
				}
			}
		}

		public void GetPasive()
		{
			using (OleDbConnection connection = new OleDbConnection(_connString))
			{
				connection.Open();
				OleDbDataReader reader = null;
				OleDbCommand command = new OleDbCommand("SELECT * FROM Pasive", connection);
				reader = command.ExecuteReader();
				while (reader.Read())
				{
					listPasive.Add(new Pasive(reader[0].ToString(), int.Parse(reader[1].ToString()),
						int.Parse(reader[2].ToString()), int.Parse(reader[3].ToString()), int.Parse(reader[4].ToString()),
						int.Parse(reader[5].ToString()), reader[6].ToString()));
				}
			}
		}

		public void GetBilant()
		{
			using (OleDbConnection connection = new OleDbConnection(_connString))
			{
				connection.Open();
				OleDbDataReader reader = null;
				OleDbCommand command = new OleDbCommand("SELECT * FROM Bilant", connection);
				reader = command.ExecuteReader();
				while (reader.Read())
				{
					listBilant.Add(new Bilant(int.Parse(reader[0].ToString()), reader[1].ToString(),
						int.Parse(reader[2].ToString()), int.Parse(reader[3].ToString()), int.Parse(reader[4].ToString()),
						int.Parse(reader[5].ToString()), int.Parse(reader[6].ToString()), int.Parse(reader[7].ToString()),
						reader[8].ToString(), int.Parse(reader[9].ToString()), int.Parse(reader[9].ToString()), int.Parse(reader[9].ToString()),
						 int.Parse(reader[9].ToString()), int.Parse(reader[9].ToString()), reader[9].ToString()));
				}
			}
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
						int.Parse(reader[5].ToString())));
				}
			}
		}

		private void Button2_Click(object sender, RibbonControlEventArgs e)
		{
			AddActivePasive addActivePasive = new AddActivePasive();
			addActivePasive.Show();
		}

		private void Button3_Click(object sender, RibbonControlEventArgs e)
		{
			ProfitCalc profitCalc = new ProfitCalc();
			profitCalc.Show();
		}
	}
}
