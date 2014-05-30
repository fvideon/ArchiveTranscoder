using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Diagnostics;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Collect additional settings for the streaming target.
	/// </summary>
	public class StreamSettingsForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBoxAsx;
		private System.Windows.Forms.GroupBox groupBoxWbv;
		private System.Windows.Forms.GroupBox groupBoxSlideImages;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox textBoxSlideBaseUrl;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.CheckBox checkBoxCreateAsx;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.TextBox textBoxWmvUrl;
		private System.Windows.Forms.TextBox textBoxPresentationUrl;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.CheckBox checkBoxCreateWbv;
		private System.Windows.Forms.TextBox textBoxAsxUrl;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Button buttonOk;
		private System.Windows.Forms.Button buttonCancel;
		private System.Windows.Forms.Button buttonSaveSettings;
		private System.Windows.Forms.Button buttonRestoreDefault;

		private System.ComponentModel.Container components = null;

		public bool CreateAsx
		{
			get { return this.checkBoxCreateAsx.Checked; }
		}
		public bool CreateWbv
		{
			get { return this.checkBoxCreateWbv.Checked; }
		}
		public String SlideBaseUrl
		{
			get { return this.textBoxSlideBaseUrl.Text.Trim(); }
		}
		public String WmvUrl
		{
			get { return this.textBoxWmvUrl.Text.Trim(); }
		}
		public String PresentationUrl
		{
			get { return this.textBoxPresentationUrl.Text.Trim(); }
		}
		public String AsxUrl
		{
			get { return this.textBoxAsxUrl.Text.Trim(); }
		}


		public StreamSettingsForm()
		{
			InitializeComponent();

			this.restoreRegSettings();
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(StreamSettingsForm));
            this.label1 = new System.Windows.Forms.Label();
            this.groupBoxAsx = new System.Windows.Forms.GroupBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.textBoxPresentationUrl = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.textBoxWmvUrl = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.checkBoxCreateAsx = new System.Windows.Forms.CheckBox();
            this.groupBoxWbv = new System.Windows.Forms.GroupBox();
            this.textBoxAsxUrl = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.checkBoxCreateWbv = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBoxSlideImages = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxSlideBaseUrl = new System.Windows.Forms.TextBox();
            this.buttonOk = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonSaveSettings = new System.Windows.Forms.Button();
            this.buttonRestoreDefault = new System.Windows.Forms.Button();
            this.groupBoxAsx.SuspendLayout();
            this.groupBoxWbv.SuspendLayout();
            this.groupBoxSlideImages.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(16, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(552, 51);
            this.label1.TabIndex = 0;
            this.label1.Text = resources.GetString("label1.Text");
            // 
            // groupBoxAsx
            // 
            this.groupBoxAsx.Controls.Add(this.label8);
            this.groupBoxAsx.Controls.Add(this.label7);
            this.groupBoxAsx.Controls.Add(this.textBoxPresentationUrl);
            this.groupBoxAsx.Controls.Add(this.label6);
            this.groupBoxAsx.Controls.Add(this.textBoxWmvUrl);
            this.groupBoxAsx.Controls.Add(this.label5);
            this.groupBoxAsx.Controls.Add(this.label2);
            this.groupBoxAsx.Controls.Add(this.checkBoxCreateAsx);
            this.groupBoxAsx.Location = new System.Drawing.Point(16, 136);
            this.groupBoxAsx.Name = "groupBoxAsx";
            this.groupBoxAsx.Size = new System.Drawing.Size(584, 264);
            this.groupBoxAsx.TabIndex = 4;
            this.groupBoxAsx.TabStop = false;
            this.groupBoxAsx.Text = "ASX File";
            // 
            // label8
            // 
            this.label8.Location = new System.Drawing.Point(16, 184);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(560, 24);
            this.label8.TabIndex = 12;
            this.label8.Text = "If the job includes presentation data, it will be packaged for WebViewer in a XML" +
                " file.   The Presentation Data URL is the HTTP location where that file will be " +
                "available.";
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(16, 98);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(560, 40);
            this.label7.TabIndex = 11;
            this.label7.Text = resources.GetString("label7.Text");
            // 
            // textBoxPresentationUrl
            // 
            this.textBoxPresentationUrl.Enabled = false;
            this.textBoxPresentationUrl.Location = new System.Drawing.Point(136, 224);
            this.textBoxPresentationUrl.Name = "textBoxPresentationUrl";
            this.textBoxPresentationUrl.Size = new System.Drawing.Size(432, 20);
            this.textBoxPresentationUrl.TabIndex = 10;
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(2, 224);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(128, 23);
            this.label6.TabIndex = 9;
            this.label6.Text = "Presentation Data URL:";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // textBoxWmvUrl
            // 
            this.textBoxWmvUrl.Enabled = false;
            this.textBoxWmvUrl.Location = new System.Drawing.Point(136, 144);
            this.textBoxWmvUrl.Name = "textBoxWmvUrl";
            this.textBoxWmvUrl.Size = new System.Drawing.Size(432, 20);
            this.textBoxWmvUrl.TabIndex = 8;
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(66, 144);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(64, 23);
            this.label5.TabIndex = 7;
            this.label5.Text = "WMV URL:";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(16, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(560, 56);
            this.label2.TabIndex = 0;
            this.label2.Text = resources.GetString("label2.Text");
            // 
            // checkBoxCreateAsx
            // 
            this.checkBoxCreateAsx.Location = new System.Drawing.Point(16, 72);
            this.checkBoxCreateAsx.Name = "checkBoxCreateAsx";
            this.checkBoxCreateAsx.Size = new System.Drawing.Size(104, 24);
            this.checkBoxCreateAsx.TabIndex = 6;
            this.checkBoxCreateAsx.Text = "Create ASX File";
            this.checkBoxCreateAsx.CheckedChanged += new System.EventHandler(this.checkBoxCreateAsx_CheckedChanged);
            // 
            // groupBoxWbv
            // 
            this.groupBoxWbv.Controls.Add(this.textBoxAsxUrl);
            this.groupBoxWbv.Controls.Add(this.label9);
            this.groupBoxWbv.Controls.Add(this.checkBoxCreateWbv);
            this.groupBoxWbv.Controls.Add(this.label3);
            this.groupBoxWbv.Enabled = false;
            this.groupBoxWbv.Location = new System.Drawing.Point(16, 408);
            this.groupBoxWbv.Name = "groupBoxWbv";
            this.groupBoxWbv.Size = new System.Drawing.Size(584, 136);
            this.groupBoxWbv.TabIndex = 5;
            this.groupBoxWbv.TabStop = false;
            this.groupBoxWbv.Text = "WBV File";
            // 
            // textBoxAsxUrl
            // 
            this.textBoxAsxUrl.Enabled = false;
            this.textBoxAsxUrl.Location = new System.Drawing.Point(143, 96);
            this.textBoxAsxUrl.Name = "textBoxAsxUrl";
            this.textBoxAsxUrl.Size = new System.Drawing.Size(432, 20);
            this.textBoxAsxUrl.TabIndex = 12;
            // 
            // label9
            // 
            this.label9.Location = new System.Drawing.Point(9, 96);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(128, 23);
            this.label9.TabIndex = 11;
            this.label9.Text = "ASX URL:";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // checkBoxCreateWbv
            // 
            this.checkBoxCreateWbv.Location = new System.Drawing.Point(16, 64);
            this.checkBoxCreateWbv.Name = "checkBoxCreateWbv";
            this.checkBoxCreateWbv.Size = new System.Drawing.Size(120, 24);
            this.checkBoxCreateWbv.TabIndex = 7;
            this.checkBoxCreateWbv.Text = "Create WBV File";
            this.checkBoxCreateWbv.CheckedChanged += new System.EventHandler(this.checkBoxCreateWbv_CheckedChanged);
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(16, 16);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(549, 40);
            this.label3.TabIndex = 0;
            this.label3.Text = resources.GetString("label3.Text");
            // 
            // groupBoxSlideImages
            // 
            this.groupBoxSlideImages.Controls.Add(this.label4);
            this.groupBoxSlideImages.Controls.Add(this.textBoxSlideBaseUrl);
            this.groupBoxSlideImages.Controls.Add(this.label1);
            this.groupBoxSlideImages.Location = new System.Drawing.Point(16, 16);
            this.groupBoxSlideImages.Name = "groupBoxSlideImages";
            this.groupBoxSlideImages.Size = new System.Drawing.Size(584, 112);
            this.groupBoxSlideImages.TabIndex = 5;
            this.groupBoxSlideImages.TabStop = false;
            this.groupBoxSlideImages.Text = "Slide Images";
            // 
            // label4
            // 
            this.label4.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.label4.Location = new System.Drawing.Point(4, 80);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(128, 23);
            this.label4.TabIndex = 2;
            this.label4.Text = "Slide Image Base URL:";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // textBoxSlideBaseUrl
            // 
            this.textBoxSlideBaseUrl.Location = new System.Drawing.Point(136, 80);
            this.textBoxSlideBaseUrl.Name = "textBoxSlideBaseUrl";
            this.textBoxSlideBaseUrl.Size = new System.Drawing.Size(432, 20);
            this.textBoxSlideBaseUrl.TabIndex = 1;
            // 
            // buttonOk
            // 
            this.buttonOk.Location = new System.Drawing.Point(528, 552);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(75, 23);
            this.buttonOk.TabIndex = 6;
            this.buttonOk.Text = "OK";
            this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(440, 552);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 7;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // buttonSaveSettings
            // 
            this.buttonSaveSettings.Location = new System.Drawing.Point(16, 552);
            this.buttonSaveSettings.Name = "buttonSaveSettings";
            this.buttonSaveSettings.Size = new System.Drawing.Size(184, 24);
            this.buttonSaveSettings.TabIndex = 8;
            this.buttonSaveSettings.Text = "Save These Settings as Default";
            this.buttonSaveSettings.Click += new System.EventHandler(this.buttonSaveSettings_Click);
            // 
            // buttonRestoreDefault
            // 
            this.buttonRestoreDefault.Location = new System.Drawing.Point(208, 552);
            this.buttonRestoreDefault.Name = "buttonRestoreDefault";
            this.buttonRestoreDefault.Size = new System.Drawing.Size(136, 24);
            this.buttonRestoreDefault.TabIndex = 9;
            this.buttonRestoreDefault.Text = "Restore Default Settings";
            this.buttonRestoreDefault.Click += new System.EventHandler(this.buttonRestoreDefault_Click);
            // 
            // StreamSettingsForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(618, 592);
            this.Controls.Add(this.buttonRestoreDefault);
            this.Controls.Add(this.buttonSaveSettings);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOk);
            this.Controls.Add(this.groupBoxAsx);
            this.Controls.Add(this.groupBoxWbv);
            this.Controls.Add(this.groupBoxSlideImages);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "StreamSettingsForm";
            this.ShowInTaskbar = false;
            this.Text = "Settings for Streaming Target";
            this.groupBoxAsx.ResumeLayout(false);
            this.groupBoxAsx.PerformLayout();
            this.groupBoxWbv.ResumeLayout(false);
            this.groupBoxWbv.PerformLayout();
            this.groupBoxSlideImages.ResumeLayout(false);
            this.groupBoxSlideImages.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// Get results
		/// </summary>
		/// <returns></returns>
		public ArchiveTranscoderJobTarget GetJobTarget()
		{
			ArchiveTranscoderJobTarget target = new ArchiveTranscoderJobTarget();
			target.Type = "stream";
			target.AsxUrl = this.textBoxAsxUrl.Text.Trim();
			target.CreateAsx = this.checkBoxCreateAsx.Checked.ToString();
            if (target.CreateAsx.ToLower().Equals("false")) {
                target.CreateWbv = "false";
            }
            else {
                target.CreateWbv = this.checkBoxCreateWbv.Checked.ToString();
            }
            string purl = this.textBoxPresentationUrl.Text.Trim();
            target.PresentationUrl = purl.Equals(string.Empty) ? null : purl;
			string sburl = this.textBoxSlideBaseUrl.Text.Trim();
            target.SlideBaseUrl = sburl.Equals(string.Empty) ? null : sburl;
			target.WmvUrl = this.textBoxWmvUrl.Text.Trim();
			return target;
		}

		/// <summary>
		/// Set existing job values.
		/// </summary>
		/// <param name="target"></param>
		public void SetJobTarget(ArchiveTranscoderJobTarget target)
		{
			this.textBoxAsxUrl.Text = target.AsxUrl;
			this.textBoxPresentationUrl.Text = target.PresentationUrl;
			this.textBoxSlideBaseUrl.Text = target.SlideBaseUrl;
			this.textBoxWmvUrl.Text = target.WmvUrl;
			this.checkBoxCreateAsx.Checked = Convert.ToBoolean(target.CreateAsx);
			this.checkBoxCreateWbv.Checked = Convert.ToBoolean(target.CreateWbv);
		}
		
		private void buttonOk_Click(object sender, System.EventArgs e)
		{
			String err = validateAll();
			if (err != null)
			{
				MessageBox.Show(err);
				return;
			}

			this.DialogResult = DialogResult.OK;
			this.Close();
		}

		private String validateAll()
		{
			if (this.textBoxSlideBaseUrl.Text.Trim() != "")
			{
				if (! validateUri(this.textBoxSlideBaseUrl.Text))
				{
					return "If Slide Image Base URL is specified, it must be in the form of a HTTP URL.";
				}
			}

			if (this.checkBoxCreateAsx.Checked)
			{
				if (this.textBoxWmvUrl.Text.Trim() == "")
				{
					return "To create the ASX file, you must specify at least the WMV URL.";
				}
				if (! validateUri(this.textBoxWmvUrl.Text))
				{
					return "The WMV URL must be in the form of a MMS or HTTP URL.";
				}

				if (this.textBoxPresentationUrl.Text.Trim() != "")
				{
					if (! validateUri(this.textBoxPresentationUrl.Text))
					{
						return "If Presentation Data URL is specified, it must be in the form of a HTTP URL.";
					}
				}
			}

			if (this.checkBoxCreateWbv.Checked)
			{
				if (this.textBoxAsxUrl.Text.Trim() == "")
				{
					return "To create a WBV file, you must specify the ASX URL.";
				}			
				if (! validateUri(this.textBoxAsxUrl.Text))
				{
					return "The ASX URL must be in the form of a HTTP URL.";
				}
			}

			return null;
		}

		/// <summary>
		/// Test that a string can be sucessfully parsed as a URI.
		/// </summary>
		/// <param name="uri"></param>
		/// <returns></returns>
		private bool validateUri(String uri)
		{
			try
			{
				Uri u = new Uri(uri.Trim());
			}
			catch
			{
				return false;
			}
			return true;
		}

		private void buttonCancel_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.Cancel;
			this.Close();
		}

		private void checkBoxCreateAsx_CheckedChanged(object sender, System.EventArgs e)
		{
			this.textBoxWmvUrl.Enabled = this.checkBoxCreateAsx.Checked;
			this.textBoxPresentationUrl.Enabled = this.checkBoxCreateAsx.Checked;
			this.groupBoxWbv.Enabled = this.checkBoxCreateAsx.Checked;
		}

		private void checkBoxCreateWbv_CheckedChanged(object sender, System.EventArgs e)
		{
			this.textBoxAsxUrl.Enabled = this.checkBoxCreateWbv.Checked;
		}

		private void restoreRegSettings()
		{
			bool createWbv = false;
			bool createAsx = false;

			try
			{
				RegistryKey BaseKey = Registry.CurrentUser.OpenSubKey(Constants.AppRegKey, true);
				if ( BaseKey == null) 
				{ //no configuration yet.. first run.
					//Debug.WriteLine("No registry configuration found.");
				}
				else
				{
					this.textBoxSlideBaseUrl.Text = Convert.ToString(BaseKey.GetValue("SlideBaseUrl",""));
					this.textBoxWmvUrl.Text = Convert.ToString(BaseKey.GetValue("WmvUrl",""));
					this.textBoxPresentationUrl.Text = Convert.ToString(BaseKey.GetValue("PresentationUrl",""));
					this.textBoxAsxUrl.Text = Convert.ToString(BaseKey.GetValue("AsxUrl",""));
					createAsx = Convert.ToBoolean(BaseKey.GetValue("CreateAsx","false"));
					createWbv = Convert.ToBoolean(BaseKey.GetValue("CreateWbv","false"));
				}
			}
			catch (Exception e)
			{
				Debug.WriteLine("Exception while reading registry: " + e.ToString());
			}

			if (createAsx)
				this.checkBoxCreateAsx.Checked = true;

			if (createWbv)
				this.checkBoxCreateWbv.Checked = true;

		}

		private void saveRegSettings()
		{
			try
			{
				RegistryKey BaseKey = Registry.CurrentUser.OpenSubKey(Constants.AppRegKey, true);
				if ( BaseKey == null) 
				{
					BaseKey = Registry.CurrentUser.CreateSubKey(Constants.AppRegKey);
				}
				BaseKey.SetValue("SlideBaseUrl",this.textBoxSlideBaseUrl.Text.Trim());
				BaseKey.SetValue("WmvUrl",this.textBoxWmvUrl.Text.Trim());
				BaseKey.SetValue("PresentationUrl",this.textBoxPresentationUrl.Text.Trim());
				BaseKey.SetValue("AsxUrl",this.textBoxAsxUrl.Text.Trim());
				BaseKey.SetValue("CreateAsx",this.checkBoxCreateAsx.Checked.ToString());
				BaseKey.SetValue("CreateWbv",this.checkBoxCreateWbv.Checked.ToString());
			}
			catch
			{
				Debug.WriteLine("Exception while saving configuration.");
			}
		}

		private void buttonSaveSettings_Click(object sender, System.EventArgs e)
		{
			this.saveRegSettings();
		}

		private void buttonRestoreDefault_Click(object sender, System.EventArgs e)
		{
			this.restoreRegSettings();
		}

	}
}
