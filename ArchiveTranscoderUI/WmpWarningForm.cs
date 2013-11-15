using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Issue Windows Media Player version warning.
	/// </summary>
	public class WmpWarningForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.CheckBox checkBoxDisableWarning;
		private System.Windows.Forms.Button buttonOk;

		private System.ComponentModel.Container components = null;

		public WmpWarningForm()
		{
			InitializeComponent();
		}

		/// <summary>
		/// True if the user has ticked the checkbox to disable this warning.
		/// </summary>
		public bool DisableWarning
		{
			get { return this.checkBoxDisableWarning.Checked; }
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
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
            this.checkBoxDisableWarning = new System.Windows.Forms.CheckBox();
            this.buttonOk = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(16, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(304, 56);
            this.label1.TabIndex = 0;
            this.label1.Text = "Windows Media Player version 10 or later is recommended on the Archive Transcoder" +
                " system.   The version currently installed may may not work correctly.  See Wind" +
                "ows Update for an upgrade.";
            // 
            // checkBoxDisableWarning
            // 
            this.checkBoxDisableWarning.Location = new System.Drawing.Point(16, 88);
            this.checkBoxDisableWarning.Name = "checkBoxDisableWarning";
            this.checkBoxDisableWarning.Size = new System.Drawing.Size(200, 32);
            this.checkBoxDisableWarning.TabIndex = 1;
            this.checkBoxDisableWarning.Text = "Don\'t show me this warning again.";
            // 
            // buttonOk
            // 
            this.buttonOk.Location = new System.Drawing.Point(240, 96);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(75, 23);
            this.buttonOk.TabIndex = 2;
            this.buttonOk.Text = "OK";
            this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
            // 
            // WmpWarningForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(330, 134);
            this.Controls.Add(this.buttonOk);
            this.Controls.Add(this.checkBoxDisableWarning);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "WmpWarningForm";
            this.ShowInTaskbar = false;
            this.Text = "Windows Media Player Version Warning";
            this.ResumeLayout(false);

		}
		#endregion

		private void buttonOk_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
			this.Close();
		}
	}
}
