using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Win32;
using System.IO;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Analyze a presentation to find out which decks are referenced, then allow the user to associate
	/// them with PPT(X) or CSD or CP3 files.
	/// </summary>
	public class DeckMatcherForm : System.Windows.Forms.Form
	{
		#region Members

		private System.Windows.Forms.Button buttonAnalyze;
		private System.Windows.Forms.Button buttonAutoMatch;
		private System.Windows.Forms.Button buttonMatchDeck;
		private System.Windows.Forms.ListBox listBoxDecks;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button buttonOk;
		private System.Windows.Forms.Button buttonCancel;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.TextBox textBoxSearchDir;
		private System.Windows.Forms.Button buttonBrowse;
		private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
		private System.Windows.Forms.Button buttonRemove;
		private System.Windows.Forms.Label labelStatus;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;

		private System.ComponentModel.Container components = null;

		private String searchDir;
		private String deckDir;
		private String cname;
		private String start;
		private String end;
		private DeckMatcher deckMatcher;
		private bool autoMatching;
		private bool analyzing;
		private bool pptInstalled;

		#endregion Members

		#region Construct/Dispose ...
		public DeckMatcherForm(String cname, String payload, String start, String end, Deck[] decks)
		{
			InitializeComponent();

			pptInstalled = Utility.CheckPptIsInstalled(); //Don't match PPT if it is not installed
			this.labelStatus.Text = "Status: Idle";
			restoreRegSettings();
			this.textBoxSearchDir.Text = searchDir;

			this.cname = cname;
			this.start = start;
			this.end = end;

			deckMatcher = new DeckMatcher(cname, payload, start, end, pptInstalled);
			deckMatcher.OnAnalyzeCompleted += new DeckMatcher.analyzeCompletedHandler(deckMatcher_OnAnalyzeCompleted);
			deckMatcher.OnStatusReport += new DeckMatcher.statusReportHandler(OnStatusReport);
			deckMatcher.OnAutoMatchCompleted += new DeckMatcher.autoMatchCompletedHandler(deckMatcher_OnAutoMatchCompleted);
			deckMatcher.OnDeckFound += new DeckMatcher.deckFoundHandler(OnDeckFound);

			if (decks != null)
			{
				for (int i=0; i< decks.Length; i++)
				{
					deckMatcher.Add(decks[i]);
				}
				foreach(Deck d in deckMatcher.Decks.Values)
				{
					this.listBoxDecks.Items.Add(d);
				}
				if (this.listBoxDecks.Items.Count > 0)
				{
					this.buttonAutoMatch.Enabled = true;
				}
			}
		}

		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (deckMatcher != null)
				{
					deckMatcher.StopThreads();
					deckMatcher = null;
				}
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		private void DeckMatcherForm_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (deckMatcher!=null)
			{
				deckMatcher.StopThreads();
			}
		}

		#endregion Construct/Dispose ...

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DeckMatcherForm));
            this.buttonAnalyze = new System.Windows.Forms.Button();
            this.buttonOk = new System.Windows.Forms.Button();
            this.buttonAutoMatch = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonMatchDeck = new System.Windows.Forms.Button();
            this.listBoxDecks = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.buttonBrowse = new System.Windows.Forms.Button();
            this.textBoxSearchDir = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.buttonRemove = new System.Windows.Forms.Button();
            this.labelStatus = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonAnalyze
            // 
            this.buttonAnalyze.Location = new System.Drawing.Point(12, 80);
            this.buttonAnalyze.Name = "buttonAnalyze";
            this.buttonAnalyze.Size = new System.Drawing.Size(152, 23);
            this.buttonAnalyze.TabIndex = 0;
            this.buttonAnalyze.Text = "Analyze Presentation Data";
            this.buttonAnalyze.Click += new System.EventHandler(this.buttonAnalyze_Click);
            // 
            // buttonOk
            // 
            this.buttonOk.Location = new System.Drawing.Point(432, 360);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(75, 23);
            this.buttonOk.TabIndex = 1;
            this.buttonOk.Text = "OK";
            this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
            // 
            // buttonAutoMatch
            // 
            this.buttonAutoMatch.Enabled = false;
            this.buttonAutoMatch.Location = new System.Drawing.Point(14, 24);
            this.buttonAutoMatch.Name = "buttonAutoMatch";
            this.buttonAutoMatch.Size = new System.Drawing.Size(136, 23);
            this.buttonAutoMatch.TabIndex = 2;
            this.buttonAutoMatch.Text = "Auto-Match All Decks ...";
            this.buttonAutoMatch.Click += new System.EventHandler(this.buttonAutoMatch_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(336, 360);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 3;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // buttonMatchDeck
            // 
            this.buttonMatchDeck.Enabled = false;
            this.buttonMatchDeck.Location = new System.Drawing.Point(11, 232);
            this.buttonMatchDeck.Name = "buttonMatchDeck";
            this.buttonMatchDeck.Size = new System.Drawing.Size(88, 23);
            this.buttonMatchDeck.TabIndex = 4;
            this.buttonMatchDeck.Text = "Match Deck ...";
            this.buttonMatchDeck.Click += new System.EventHandler(this.buttonMatchDeck_Click);
            // 
            // listBoxDecks
            // 
            this.listBoxDecks.HorizontalScrollbar = true;
            this.listBoxDecks.Location = new System.Drawing.Point(11, 128);
            this.listBoxDecks.Name = "listBoxDecks";
            this.listBoxDecks.Size = new System.Drawing.Size(496, 95);
            this.listBoxDecks.TabIndex = 5;
            this.listBoxDecks.DoubleClick += new System.EventHandler(this.listBoxDecks_DoubleClick);
            this.listBoxDecks.SelectedIndexChanged += new System.EventHandler(this.listBoxDecks_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(11, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(501, 64);
            this.label1.TabIndex = 6;
            this.label1.Text = resources.GetString("label1.Text");
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(11, 108);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(144, 16);
            this.label2.TabIndex = 7;
            this.label2.Text = "Slide Decks Found:";
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(14, 56);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(96, 16);
            this.label4.TabIndex = 9;
            this.label4.Text = "Search Directory:";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.buttonBrowse);
            this.groupBox1.Controls.Add(this.textBoxSearchDir);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.buttonAutoMatch);
            this.groupBox1.Location = new System.Drawing.Point(10, 264);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(496, 88);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Auto-Match will recursively search a specified directory to find deck matches";
            // 
            // buttonBrowse
            // 
            this.buttonBrowse.Location = new System.Drawing.Point(432, 56);
            this.buttonBrowse.Name = "buttonBrowse";
            this.buttonBrowse.Size = new System.Drawing.Size(56, 23);
            this.buttonBrowse.TabIndex = 11;
            this.buttonBrowse.Text = "Browse";
            this.buttonBrowse.Click += new System.EventHandler(this.buttonBrowse_Click);
            // 
            // textBoxSearchDir
            // 
            this.textBoxSearchDir.Location = new System.Drawing.Point(104, 56);
            this.textBoxSearchDir.Name = "textBoxSearchDir";
            this.textBoxSearchDir.Size = new System.Drawing.Size(320, 20);
            this.textBoxSearchDir.TabIndex = 10;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "PPT files|*.ppt|PPTX files|*.pptx|CP3 files|*.cp3|CSD files|*.csd";
            // 
            // buttonRemove
            // 
            this.buttonRemove.Enabled = false;
            this.buttonRemove.Location = new System.Drawing.Point(112, 232);
            this.buttonRemove.Name = "buttonRemove";
            this.buttonRemove.Size = new System.Drawing.Size(88, 23);
            this.buttonRemove.TabIndex = 12;
            this.buttonRemove.Text = "Remove";
            this.buttonRemove.Click += new System.EventHandler(this.buttonRemove_Click);
            // 
            // labelStatus
            // 
            this.labelStatus.Location = new System.Drawing.Point(184, 85);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(304, 16);
            this.labelStatus.TabIndex = 13;
            // 
            // DeckMatcherForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(520, 392);
            this.Controls.Add(this.labelStatus);
            this.Controls.Add(this.buttonRemove);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.listBoxDecks);
            this.Controls.Add(this.buttonMatchDeck);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOk);
            this.Controls.Add(this.buttonAnalyze);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DeckMatcherForm";
            this.ShowInTaskbar = false;
            this.Text = "Match Presentation Decks";
            this.Closing += new System.ComponentModel.CancelEventHandler(this.DeckMatcherForm_Closing);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		#region Public Method

		/// <summary>
		/// Return the result or null if no decks.
		/// </summary>
		/// <returns></returns>
		public Deck[] ToSlideDecks()
		{
			if (deckMatcher.Decks.Count == 0)
				return null;

			Deck[] decks = new Deck[deckMatcher.Decks.Count];
			int i=0;
			foreach(Deck d in deckMatcher.Decks.Values)
			{
				decks[i] = d;
				i++;
			}
			return decks;
		}

		#endregion Public Method

		#region Private Methods


		private void restoreRegSettings()
		{
			searchDir = "";
			deckDir = "";
			try
			{
				RegistryKey BaseKey = Registry.CurrentUser.OpenSubKey(Constants.AppRegKey, true);
				if ( BaseKey == null) 
				{ //no configuration yet.. first run.
					//Debug.WriteLine("No registry configuration found.");
					return;
				}
				searchDir = Convert.ToString(BaseKey.GetValue("SearchDir",searchDir));
				deckDir = Convert.ToString(BaseKey.GetValue("DeckDir",deckDir));
			}
			catch (Exception e)
			{
				Debug.WriteLine("Exception while reading registry: " + e.ToString());
			}
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
				BaseKey.SetValue("SearchDir",searchDir);
				BaseKey.SetValue("DeckDir",deckDir);
			}
			catch
			{
				Debug.WriteLine("Exception while saving configuration.");
			}
		}

		private void analysisEnable(bool enable)
		{
			this.buttonAutoMatch.Enabled = enable;
			if (enable)
			{
				if (this.listBoxDecks.Items.Count == 0)
					this.buttonAutoMatch.Enabled = false;
			}
			this.buttonBrowse.Enabled = enable;
			this.buttonOk.Enabled = enable;
			this.buttonCancel.Enabled = enable;
		}

		private void autoMatchEnable(bool enable)
		{
			this.buttonAnalyze.Enabled = enable;
			this.buttonBrowse.Enabled = enable;
			this.buttonOk.Enabled = enable;
			this.buttonCancel.Enabled = enable;
		}

		#endregion Private Methods

		#region UI Control Event Handlers

		private void buttonCancel_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.Cancel;
			searchDir = this.textBoxSearchDir.Text;
			saveRegSettings();

			this.Close();
		}

		private void buttonOk_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
			searchDir = this.textBoxSearchDir.Text;
			saveRegSettings();

			this.Close();
		}

		private void buttonBrowse_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.OK == folderBrowserDialog1.ShowDialog())
			{
				this.textBoxSearchDir.Text = folderBrowserDialog1.SelectedPath;
			}

		}

		private void buttonAnalyze_Click(object sender, System.EventArgs e)
		{
			if (analyzing)
			{
				deckMatcher.StopAnalyze();
				deckMatcher_OnAnalyzeCompleted();
			}
			else
			{
				this.listBoxDecks.Items.Clear();
				this.listBoxDecks.Enabled = false;
				String errorMsg = deckMatcher.Analyze();

				if (errorMsg != null)
				{
					MessageBox.Show(errorMsg);
				}
				else
				{
					analysisEnable(false);
					this.Cursor = Cursors.AppStarting;
					this.buttonAnalyze.Text = "Stop Analysis";
					analyzing = true;
				}
			}
		}

		private void listBoxDecks_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.buttonMatchDeck.Enabled = true;
			this.buttonRemove.Enabled = true;
		}

		private void buttonMatchDeck_Click(object sender, System.EventArgs e)
		{
			if (this.listBoxDecks.SelectedItem == null)
			{
				this.buttonMatchDeck.Enabled = false;
				return;
			}

			this.openFileDialog1.Title = "Select Matching Deck";
			if (pptInstalled)
			{
                this.openFileDialog1.Filter = "PPT files|*.ppt|PPTX files|*.pptx|CP3 files|*.cp3|CSD files|*.csd";
			}
			else
			{
				this.openFileDialog1.Filter = "CP3 files|*.cp3|CSD files|*.csd";		
			}
			if (deckDir != "")
			{
				openFileDialog1.InitialDirectory = deckDir;
			}
			if (DialogResult.OK == openFileDialog1.ShowDialog())
			{
				Deck d = ((Deck)this.listBoxDecks.SelectedItem);
				d.Matched = true;
				d.Path = openFileDialog1.FileName;
				this.listBoxDecks.Items.Remove(this.listBoxDecks.SelectedItem);
				this.listBoxDecks.Items.Add(d);
				this.buttonMatchDeck.Enabled = false;
				this.buttonRemove.Enabled = false;
				deckDir = d.Path;
			}
		}

		private void listBoxDecks_DoubleClick(object sender, System.EventArgs e)
		{
			if (this.buttonMatchDeck.Enabled)
				this.buttonMatchDeck_Click(sender, e);
		}

		private void buttonAutoMatch_Click(object sender, System.EventArgs e)
		{
			if (autoMatching)
			{
				deckMatcher.StopAutoMatch();
				deckMatcher_OnAutoMatchCompleted();
			}
			else
			{
				searchDir = this.textBoxSearchDir.Text.Trim();
				if (searchDir == "")
				{
					MessageBox.Show("First specify the search directory.");
					return;
				}
				if (!Directory.Exists(searchDir))
				{
					MessageBox.Show("Specified search directory does not exist: " + searchDir);
					return;
				}
				if (this.listBoxDecks.Items.Count == 0)
				{
					MessageBox.Show("There are no decks to match.");
					this.buttonAutoMatch.Enabled = false;
					return;
				}

				autoMatching = true;
				this.Cursor = Cursors.AppStarting;
				this.autoMatchEnable(false);
				this.buttonAutoMatch.Text = "Cancel Auto-Match";
				deckMatcher.AutoMatch(new DirectoryInfo(searchDir));
			}
		}


		private void buttonRemove_Click(object sender, System.EventArgs e)
		{
			if (this.listBoxDecks.SelectedItem != null)
			{
				deckMatcher.Remove((Deck)this.listBoxDecks.SelectedItem);
				this.listBoxDecks.Items.Remove(this.listBoxDecks.SelectedItem);
			}
			this.buttonRemove.Enabled=false;
			this.buttonMatchDeck.Enabled=false;
		}

		#endregion UI Control Event Handlers

		#region DeckMatcher Event Handlers

		private void OnStatusReport(String message)
		{
			object[] oa = new object[1];
			oa[0] = message;
			if (this.InvokeRequired)
			{
				this.Invoke(new StatusReportDelegate(StatusReport),oa);
			}
			else
			{
				StatusReport(message);
			}
		}

		private delegate void StatusReportDelegate(String message);
		private void StatusReport(String message)
		{
			this.labelStatus.Text = message;
		}

		private void OnDeckFound(Deck deck)
		{
			object[] oa = new object[1];
			oa[0] = deck;
			if (this.InvokeRequired)
			{
				this.Invoke(new addDeckDelegate(addDeck),oa);
			}
			else
			{
				addDeck(deck);
			}
		}

		private delegate void addDeckDelegate(Deck deck);
		private void addDeck(Deck deck)
		{
			this.listBoxDecks.Items.Add(deck);
		}

		private void deckMatcher_OnAnalyzeCompleted()
		{
			if (this.InvokeRequired)
			{
				this.Invoke(new AnalyzeCompletedDelegate(AnalyzeCompleted));
			}
			else
			{
				AnalyzeCompleted();
			}
		}

		private delegate void AnalyzeCompletedDelegate();

		private void AnalyzeCompleted()
		{
			this.listBoxDecks.Enabled = true;
			this.buttonMatchDeck.Enabled = false;
			this.buttonRemove.Enabled = false;

			if (this.listBoxDecks.Items.Count > 0)
			{
				this.buttonAutoMatch.Enabled = true;
				this.listBoxDecks.SelectedItem = this.listBoxDecks.Items[0];
			}

			analysisEnable(true);
			this.buttonAnalyze.Text = "Analyze Presentation Data";
			this.labelStatus.Text = "Status: Idle";
			this.Cursor = Cursors.Default;
			analyzing = false;
		}

		private void deckMatcher_OnAutoMatchCompleted()
		{
			if (this.InvokeRequired)
			{
				this.Invoke(new AutoMatchCompletedDelegate(AutoMatchCompleted));
			}
			else
			{
				AutoMatchCompleted();
			}
		}

		private delegate void AutoMatchCompletedDelegate();

		private void AutoMatchCompleted()
		{
			this.listBoxDecks.Items.Clear();
			foreach(Deck d in deckMatcher.Decks.Values)
			{
				this.listBoxDecks.Items.Add(d);
				this.buttonMatchDeck.Enabled = false;
				this.buttonRemove.Enabled = false;
			}
			this.buttonAutoMatch.Text = "Auto-Match All Decks ...";
			autoMatchEnable(true);
			autoMatching = false;
			this.Cursor = Cursors.Default;
		}

		#endregion DeckMatcher Event Handlers

	}
}
