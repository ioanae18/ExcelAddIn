namespace ExcelAddIn
{
	partial class ProfitCalc
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.btnShow = new System.Windows.Forms.Button();
			this.txtBoxCost = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// btnShow
			// 
			this.btnShow.Location = new System.Drawing.Point(131, 110);
			this.btnShow.Name = "btnShow";
			this.btnShow.Size = new System.Drawing.Size(75, 23);
			this.btnShow.TabIndex = 0;
			this.btnShow.Text = "Show";
			this.btnShow.UseVisualStyleBackColor = true;
			this.btnShow.Click += new System.EventHandler(this.BtnShow_Click);
			// 
			// txtBoxCost
			// 
			this.txtBoxCost.Location = new System.Drawing.Point(103, 38);
			this.txtBoxCost.Name = "txtBoxCost";
			this.txtBoxCost.Size = new System.Drawing.Size(146, 20);
			this.txtBoxCost.TabIndex = 1;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(22, 41);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 13);
			this.label1.TabIndex = 2;
			this.label1.Text = "Costs";
			// 
			// ProfitCalc
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(800, 450);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtBoxCost);
			this.Controls.Add(this.btnShow);
			this.Name = "ProfitCalc";
			this.Text = "ProfitCalc";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button btnShow;
		private System.Windows.Forms.TextBox txtBoxCost;
		private System.Windows.Forms.Label label1;
	}
}