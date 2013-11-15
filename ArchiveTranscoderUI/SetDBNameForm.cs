using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Simple input dialog to let the user set a non-standard DB name
	/// </summary>
	public class SetDBNameForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button buttonOk;
		private System.Windows.Forms.Button buttonCancel;
		private System.Windows.Forms.TextBox textBoxDBName;
		private System.Windows.Forms.Label label2;
		private System.ComponentModel.Container components = null;

		public String DBName
		{
			get
			{
				if (this.textBoxDBName.Text.Trim() != "")
					return this.textBoxDBName.Text.Trim();
				else
					return "ArchiveService";
			}
			set
			{
				this.textBoxDBName.Text = value;
			}
		}
		public SetDBNameForm()
		{
			InitializeComponent();
		}

		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.label1 = new System.Windows.Forms.Label();
            this.buttonOk = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.textBoxDBName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(8, 64);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "Database Name:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // buttonOk
            // 
            this.buttonOk.Location = new System.Drawing.Point(304, 104);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(75, 23);
            this.buttonOk.TabIndex = 1;
            this.buttonOk.Text = "OK";
            this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(208, 104);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 2;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // textBoxDBName
            // 
            this.textBoxDBName.Location = new System.Drawing.Point(112, 64);
            this.textBoxDBName.Name = "textBoxDBName";
            this.textBoxDBName.Size = new System.Drawing.Size(264, 20);
            this.textBoxDBName.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(16, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(360, 32);
            this.label2.TabIndex = 4;
            this.label2.Text = "The Archive Service database is normally named \'ArchiveService\'.  To operate on a" +
                " database with a different name, set the name here.";
            // 
            // SetDBNameForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(394, 144);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBoxDBName);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOk);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SetDBNameForm";
            this.ShowInTaskbar = false;
            this.Text = "Set ArchiveService Database Name";
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		private void buttonOk_Click(object sender, System.EventArgs e)
		{
			if (this.textBoxDBName.Text.Trim() != "")
			{
				this.DialogResult = DialogResult.OK;
				this.Close();
			} 
			else
			{
				MessageBox.Show("You must specify a database name.  'ArchiveService' is the default name.");
			}
		}

		private void buttonCancel_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.Cancel;
			this.Close();
		}
	}
}
