using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Windows.Forms;
using DBModels;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
	public partial class AddActivePasive : Form
	{
		private const string _connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\Faculta\FACULTA_MASTER\ANULI-SEM1\Ingineria Programarii si Limbaje de Asamblare\ProiectFinalIP\ExcelAddIn\MagazinDB.accdb";

		private List<Active> listActive = new List<Active>();
		private List<Pasive> listPasive = new List<Pasive>();
		private List<Store> listStore = new List<Store>();

		public AddActivePasive()
		{
			InitializeComponent();
			AddCmb();
		}

		private void AddCmb()
		{
			GetActive();
			GetPasive();
			GetStore();

			foreach(var a in listActive)
			{
				a.store = listStore.Where(x => x.StoreID == a.ActiveID).FirstOrDefault();
			}		

			cmbStore.DataSource = listStore.Select(x => new { id = x.StoreID, name = x.StoreName }).ToList();
			cmbStore.ValueMember = "id";
			cmbStore.DisplayMember = "name";
			cmbNec.DataSource = listActive.Select(x => new { id = x.ImobNecorp }).ToList();
			cmbNec.ValueMember = "id";
			cmbCorp.DataSource = listActive.Select(x => new { id = x.ImobCorp }).ToList();
			cmbCorp.ValueMember = "id";
			cmbFin.DataSource = listActive.Select(x => new { id = x.ImobFin }).ToList();
			cmbFin.ValueMember = "id";
			cmbStoc.DataSource = listActive.Select(x => new { id = x.Stocuri }).ToList();
			cmbStoc.ValueMember = "id";
			cmbCrean.DataSource = listActive.Select(x => new { id = x.Creante }).ToList();
			cmbCrean.ValueMember = "id";
			cmbCCB.DataSource = listActive.Select(x => new { id = x.CasaConturi }).ToList();
			cmbCCB.ValueMember = "id";
			cmbProp.DataSource = listPasive.Select(x => new { id = x.CapProprii }).ToList();
			cmbProp.ValueMember = "id";
			cmbProv.DataSource = listPasive.Select(x => new { id = x.Provizioane }).ToList();
			cmbProv.ValueMember = "id";
			cmbDatL.DataSource = listPasive.Select(x => new { id = x.DatLung }).ToList();
			cmbDatL.ValueMember = "id";
			cmbDatS.DataSource = listPasive.Select(x => new { id = x.DatScurt }).ToList();
			cmbDatS.ValueMember = "id";
			cmbAvans.DataSource = listPasive.Select(x => new { id = x.VenitAvans }).ToList();
			cmbAvans.ValueMember = "id";
		}

		private void GetActive()
		{
			try
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

			catch
			{
				MessageBox.Show("Database connection error!");
			}
		}

		public void GetPasive()
		{
			try
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
			catch
			{
				MessageBox.Show("Database connection error!");
			}
		}

		public void GetStore()
		{
			try
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
			catch
			{
				MessageBox.Show("Database connection error!");
			}
		}

		private void BtnSave_Click(object sender, EventArgs e)
		{
			try
			{
				Bilant bil = new DBModels.Bilant();
				bil.BilantID = int.Parse(txtBilID.Text);
				bil.ImobNecorp = int.Parse(cmbNec.Text);
				bil.ImobCorp = int.Parse(cmbCorp.Text);
				bil.ImobFin = int.Parse(cmbFin.Text);
				bil.Stocuri = int.Parse(cmbStoc.Text);
				bil.Creante = int.Parse(cmbCrean.Text);
				bil.CasaConturi = int.Parse(cmbCCB.Text);
				bil.CapProprii = int.Parse(cmbProp.Text);
				bil.Provizioane = int.Parse(cmbProv.Text);
				bil.DatLung = int.Parse(cmbDatL.Text);
				bil.DatScurt = int.Parse(cmbDatS.Text);
				bil.VenitAvans = int.Parse(cmbAvans.Text);
				bil.StoreID = cmbStore.SelectedValue.ToString();

				using (OleDbConnection connection = new OleDbConnection(_connString))
				{
					connection.Open();
					OleDbCommand command = new OleDbCommand(@"
INSERT INTO Bilant (BilantID, ImobNecorp, ImobCorp, ImobFin, Stocuri, Creante, CasaConturi, CapProprii, Provizioane, DatLung, DatScurt, VenitAvans, StoreID) 
VALUES (@BilantID, @ImobNecorp, @ImobCorp, @ImobFin, @Stocuri, @Creante, @CasaConturi, @CapProprii, @Provizioane, @DatLung, @DatScurt, @VenitAvans, @StoreID)", connection);
					command.Parameters.AddWithValue("@BilantID", bil.BilantID);
					command.Parameters.AddWithValue("@ImobNecorp", bil.ImobNecorp);
					command.Parameters.AddWithValue("@ImobCorp", bil.ImobCorp);
					command.Parameters.AddWithValue("@ImobFin", bil.ImobFin);
					command.Parameters.AddWithValue("@Stocuri", bil.Stocuri);
					command.Parameters.AddWithValue("@Creante", bil.Creante);
					command.Parameters.AddWithValue("@CasaConturi", bil.CasaConturi);
					command.Parameters.AddWithValue("@CapProprii", bil.CapProprii);
					command.Parameters.AddWithValue("@Provizioane", bil.Provizioane);
					command.Parameters.AddWithValue("@DatLung", bil.DatLung);
					command.Parameters.AddWithValue("@DatScurt", bil.DatScurt);
					command.Parameters.AddWithValue("@VenitAvans", bil.VenitAvans);
					command.Parameters.AddWithValue("@StoreID", bil.StoreID);

					var savedRows = command.ExecuteNonQuery();
				}
				MessageBox.Show("Saved!");
			}
			catch
			{
				MessageBox.Show("Server error!");
			}
		}

		private void BtnClose_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void BtnGen_Click(object sender, EventArgs e)
		{
			//create Word file into template and fill data
			Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
			//load document
			Microsoft.Office.Interop.Word.Document doc = null;
			object fileName = "E:\\Faculta\\FACULTA_MASTER\\ANULI-SEM1\\Ingineria Programarii si Limbaje de Asamblare\\ProiectFinalIP\\Template.docx";
			object missing = Type.Missing;

			for ( int i = 0; i < 13; i ++)
			{
				doc = app.Documents.Open(fileName, missing, missing);
				app.Selection.Find.ClearFormatting();
				app.Selection.Find.Replacement.ClearFormatting();

				//read excel file
				string[] tmp = new string[13];
				tmp = readExcel(i);

				//fill data to template 
				app.Selection.Find.Execute("<BilantID>", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, tmp[0]);
				app.Selection.Find.Execute("<ImobNecorp>", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, tmp[1]);
				app.Selection.Find.Execute("<ImobCorp>", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, tmp[2]);
				app.Selection.Find.Execute("<ImobFin>", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, tmp[3]);
				app.Selection.Find.Execute("<Stocuri>", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, tmp[4]);
				app.Selection.Find.Execute("<Creante>", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, tmp[5]);
				app.Selection.Find.Execute("<CasaConturi>", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, tmp[6]);
				app.Selection.Find.Execute("<CapProprii>", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, tmp[7]);
				app.Selection.Find.Execute("<Provizioane>", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, tmp[8]);
				app.Selection.Find.Execute("<DatLung>", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, tmp[9]);
				app.Selection.Find.Execute("<DatScurt>", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, tmp[10]);
				app.Selection.Find.Execute("<VenitAvans>", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, tmp[11]);
				app.Selection.Find.Execute("<StoreID>", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, tmp[12]);

				//save new file as report
				object SaveAsFile = (object)"E:\\Faculta\\FACULTA_MASTER\\ANULI-SEM1\\Ingineria Programarii si Limbaje de Asamblare\\ProiectFinalIP\\Reports" + tmp[0] + ".doc";
				doc.SaveAs2(SaveAsFile, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
			}
			MessageBox.Show("Files are created!");
			doc.Close(false, missing, missing);
			app.Quit(false, false, false);
			System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
		}

		private string[] readExcel(int index)
		{
			string res = "E:\\Faculta\\FACULTA_MASTER\\ANULI-SEM1\\Ingineria Programarii si Limbaje de Asamblare\\ProiectFinalIP\\Excel.xlsx";
			Excel.Application xlApp;
			Excel.Workbook xlWorkbook;
			Excel.Worksheet xlWorksheet;

			xlApp = new Excel.Application();
			xlWorkbook = xlApp.Workbooks.Open(res, 0, true, 5, "", "", true);
			xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

			index += 2;
			string[] data = new string[13]; 
			data[0] = xlWorksheet.get_Range("A" + index.ToString()).Text;
			data[1] = xlWorksheet.get_Range("B" + index.ToString()).Text;
			data[2] = xlWorksheet.get_Range("C" + index.ToString()).Text;
			data[3] = xlWorksheet.get_Range("D" + index.ToString()).Text;
			data[4] = xlWorksheet.get_Range("E" + index.ToString()).Text;
			data[5] = xlWorksheet.get_Range("F" + index.ToString()).Text;
			data[6] = xlWorksheet.get_Range("G" + index.ToString()).Text;
			data[7] = xlWorksheet.get_Range("H" + index.ToString()).Text;
			data[8] = xlWorksheet.get_Range("I" + index.ToString()).Text;
			data[9] = xlWorksheet.get_Range("J" + index.ToString()).Text;
			data[10] = xlWorksheet.get_Range("K" + index.ToString()).Text;
			data[11] = xlWorksheet.get_Range("L" + index.ToString()).Text;
			data[12] = xlWorksheet.get_Range("M" + index.ToString()).Text;

			xlWorkbook.Close(false);
			xlApp.Quit();
			System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
			System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
			System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

			return data;
		}
	}
}
