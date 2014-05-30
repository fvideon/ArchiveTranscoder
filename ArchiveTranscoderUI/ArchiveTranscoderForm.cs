using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using Microsoft.Win32;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Main form for Archive Transcoder
	/// </summary>
	public class ArchiveTranscoderForm : System.Windows.Forms.Form
	{
		#region Private Members

		private String logFile;					//Path to the log file of the last job run, or null.
		private String currentDoc;				//path of currently opened document file.
		private bool dirtyBit;					//Indicates that something in the form changed since last doc file save
		private bool exitAfterStop;				//Flag to exit app after transcoder has stopped.
		private static string launchFile;		//File name provided on the command line
		private ArchiveTranscoder transcoder;	//The main transcoder object
		private ArchiveTranscoderBatch batch;	//Settings for current jobs
		private bool encoding = false;			//Indicates transcoder is currently running
		private String jobDir;					//Default place to look for job XML files
		private bool closing = false;			//flags that the form is closing
		private ConferenceDataCache confDataCache;

		#endregion Private Members

		#region Form Designer Members

        private IContainer components;
        private System.Windows.Forms.Label label1;
		private System.Windows.Forms.MainMenu mainMenu1;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem menuItem5;
		private System.Windows.Forms.MenuItem menuItem8;
		private System.Windows.Forms.TextBox textBoxJobTitle;
		private System.Windows.Forms.ListBox listBoxSegments;
		private System.Windows.Forms.Button buttonCreateSegment;
		private System.Windows.Forms.Button buttonEditSegment;
		private System.Windows.Forms.Button buttonDeleteSegment;
		private System.Windows.Forms.Button buttonStartStop;
		private System.Windows.Forms.CheckBox checkBoxStream;
		private System.Windows.Forms.CheckBox checkBoxDownload;
		private System.Windows.Forms.MenuItem menuItemLoadJob;
		private System.Windows.Forms.MenuItem menuItemSaveJob;
		private System.Windows.Forms.MenuItem menuItemExit;
		private System.Windows.Forms.MenuItem menuItemAbout;
		private System.Windows.Forms.MenuItem menuItemRunBatch;
		private System.Windows.Forms.MenuItem menuItemSelectService;
		private System.Windows.Forms.TextBox textBoxOutputDir;
		private System.Windows.Forms.TextBox textBoxBaseName;
		private System.Windows.Forms.Button buttonBrowse;
		private System.Windows.Forms.ComboBox comboBoxProfile;
		private System.Windows.Forms.GroupBox groupBoxSegments;
		private System.Windows.Forms.GroupBox groupBoxTarget;
		private System.Windows.Forms.GroupBox groupBoxProfile;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
		private System.Windows.Forms.Label labelStatus;
		private System.Windows.Forms.Button buttonCopySegment;
		private System.Windows.Forms.MenuItem menuItemSetDB;
		private System.Windows.Forms.Button buttonStreamSettings;
		private System.Windows.Forms.CheckBox checkBox1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.MenuItem menuItemViewLog;
		private ArchiveTranscoderJobTarget streamTarget;
		private System.Windows.Forms.MenuItem menuItemNewJob;
		private System.Windows.Forms.Label labelBaseName;
		private System.Windows.Forms.Label labelOutputDir;
		private System.Windows.Forms.MenuItem menuItemSaveAs;
		#endregion Form Designer Members

		#region Construct/Dispose/Main

		public ArchiveTranscoderForm()
		{
			InitializeComponent();
			confDataCache = null;
			streamTarget = null;
			logFile = null;
			batch = null;
			currentDoc = null;
			dirtyBit = false;
			exitAfterStop = false;
			transcoder = new ArchiveTranscoder();
			transcoder.OnBatchCompleted += new ArchiveTranscoder.batchCompletedHandler(OnBatchCompleted);
			transcoder.OnStatusReport += new ArchiveTranscoder.statusReportHandler(OnStatusReport);
			transcoder.OnStopCompleted += new ArchiveTranscoder.stopCompletedHandler(StopCompletedHandler);
			this.restoreRegSettings();
			setToolTips();
			FillWMProfileList();
			this.cleanUpOldTempFiles();

			if (launchFile != "")
			{
				if (File.Exists(launchFile))
				{
					FileInfo fi = new FileInfo(launchFile);
					if (fi.Extension.ToLower() == ".xml")
					{
						jobDir = launchFile;
						this.loadJob(launchFile);
					}
					else
					{
						MessageBox.Show("ArchiveTranscoder cannot open files of type " + fi.Extension,"Error");
					}
				}
				else
				{
					MessageBox.Show("Launch file does not exist: " + launchFile, "Error");
				}
			}

			confDataCache = new ConferenceDataCache(transcoder.SQLHost,transcoder.DBName);
			confDataCache.Lookup();

			setFormTitle();

            //Known issues exist surrounding running without a supported version of WMP.  9-11 should work.
			//checkWmpVersion();
  		}

		protected override void Dispose( bool disposing )
		{
			transcoder.Stop();
			transcoder.Dispose();
			confDataCache.Dispose();
			this.saveRegSettings();
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main(string[] args) 
		{
            Application.EnableVisualStyles(); //<--Use the rounded buttons, etc. when running on XP

			launchFile = "";
			if (args.Length > 0)
			{
				launchFile = args[0];
			}

			Application.Run(new ArchiveTranscoderForm());
		}
		#endregion Construct/Dispose/Main

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ArchiveTranscoderForm));
            this.textBoxJobTitle = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.listBoxSegments = new System.Windows.Forms.ListBox();
            this.buttonCreateSegment = new System.Windows.Forms.Button();
            this.buttonEditSegment = new System.Windows.Forms.Button();
            this.buttonDeleteSegment = new System.Windows.Forms.Button();
            this.buttonStartStop = new System.Windows.Forms.Button();
            this.checkBoxStream = new System.Windows.Forms.CheckBox();
            this.checkBoxDownload = new System.Windows.Forms.CheckBox();
            this.mainMenu1 = new System.Windows.Forms.MainMenu(this.components);
            this.menuItem1 = new System.Windows.Forms.MenuItem();
            this.menuItemNewJob = new System.Windows.Forms.MenuItem();
            this.menuItemLoadJob = new System.Windows.Forms.MenuItem();
            this.menuItemSaveJob = new System.Windows.Forms.MenuItem();
            this.menuItemSaveAs = new System.Windows.Forms.MenuItem();
            this.menuItemRunBatch = new System.Windows.Forms.MenuItem();
            this.menuItemViewLog = new System.Windows.Forms.MenuItem();
            this.menuItemExit = new System.Windows.Forms.MenuItem();
            this.menuItem8 = new System.Windows.Forms.MenuItem();
            this.menuItemSelectService = new System.Windows.Forms.MenuItem();
            this.menuItemSetDB = new System.Windows.Forms.MenuItem();
            this.menuItem5 = new System.Windows.Forms.MenuItem();
            this.menuItemAbout = new System.Windows.Forms.MenuItem();
            this.groupBoxSegments = new System.Windows.Forms.GroupBox();
            this.buttonCopySegment = new System.Windows.Forms.Button();
            this.groupBoxTarget = new System.Windows.Forms.GroupBox();
            this.buttonStreamSettings = new System.Windows.Forms.Button();
            this.labelOutputDir = new System.Windows.Forms.Label();
            this.textBoxOutputDir = new System.Windows.Forms.TextBox();
            this.textBoxBaseName = new System.Windows.Forms.TextBox();
            this.labelBaseName = new System.Windows.Forms.Label();
            this.buttonBrowse = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.groupBoxProfile = new System.Windows.Forms.GroupBox();
            this.comboBoxProfile = new System.Windows.Forms.ComboBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.labelStatus = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBoxSegments.SuspendLayout();
            this.groupBoxTarget.SuspendLayout();
            this.groupBoxProfile.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBoxJobTitle
            // 
            this.textBoxJobTitle.Location = new System.Drawing.Point(72, 16);
            this.textBoxJobTitle.Name = "textBoxJobTitle";
            this.textBoxJobTitle.Size = new System.Drawing.Size(376, 20);
            this.textBoxJobTitle.TabIndex = 1;
            this.textBoxJobTitle.TextChanged += new System.EventHandler(this.textBoxJobTitle_TextChanged);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(13, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 16);
            this.label1.TabIndex = 2;
            this.label1.Text = "Job Title:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // listBoxSegments
            // 
            this.listBoxSegments.BackColor = System.Drawing.SystemColors.Window;
            this.listBoxSegments.Location = new System.Drawing.Point(24, 108);
            this.listBoxSegments.Name = "listBoxSegments";
            this.listBoxSegments.Size = new System.Drawing.Size(416, 82);
            this.listBoxSegments.TabIndex = 3;
            this.listBoxSegments.DoubleClick += new System.EventHandler(this.listBoxSegments_DoubleClick);
            this.listBoxSegments.SelectedIndexChanged += new System.EventHandler(this.listBoxSegments_SelectedIndexChanged);
            // 
            // buttonCreateSegment
            // 
            this.buttonCreateSegment.Location = new System.Drawing.Point(8, 20);
            this.buttonCreateSegment.Name = "buttonCreateSegment";
            this.buttonCreateSegment.Size = new System.Drawing.Size(56, 24);
            this.buttonCreateSegment.TabIndex = 4;
            this.buttonCreateSegment.Text = "Create";
            this.buttonCreateSegment.Click += new System.EventHandler(this.buttonCreateSegment_Click);
            // 
            // buttonEditSegment
            // 
            this.buttonEditSegment.Enabled = false;
            this.buttonEditSegment.Location = new System.Drawing.Point(72, 20);
            this.buttonEditSegment.Name = "buttonEditSegment";
            this.buttonEditSegment.Size = new System.Drawing.Size(56, 24);
            this.buttonEditSegment.TabIndex = 5;
            this.buttonEditSegment.Text = "Edit";
            this.buttonEditSegment.Click += new System.EventHandler(this.buttonEditSegment_Click);
            // 
            // buttonDeleteSegment
            // 
            this.buttonDeleteSegment.Enabled = false;
            this.buttonDeleteSegment.Location = new System.Drawing.Point(200, 20);
            this.buttonDeleteSegment.Name = "buttonDeleteSegment";
            this.buttonDeleteSegment.Size = new System.Drawing.Size(56, 24);
            this.buttonDeleteSegment.TabIndex = 6;
            this.buttonDeleteSegment.Text = "Remove";
            this.buttonDeleteSegment.Click += new System.EventHandler(this.buttonDeleteSegment_Click);
            // 
            // buttonStartStop
            // 
            this.buttonStartStop.Location = new System.Drawing.Point(360, 416);
            this.buttonStartStop.Name = "buttonStartStop";
            this.buttonStartStop.Size = new System.Drawing.Size(88, 24);
            this.buttonStartStop.TabIndex = 6;
            this.buttonStartStop.Text = "Start Job";
            this.buttonStartStop.Click += new System.EventHandler(this.buttonStartStop_Click);
            // 
            // checkBoxStream
            // 
            this.checkBoxStream.Location = new System.Drawing.Point(40, 88);
            this.checkBoxStream.Name = "checkBoxStream";
            this.checkBoxStream.Size = new System.Drawing.Size(64, 24);
            this.checkBoxStream.TabIndex = 7;
            this.checkBoxStream.Text = "Stream";
            this.checkBoxStream.CheckedChanged += new System.EventHandler(this.checkBoxStream_CheckedChanged);
            // 
            // checkBoxDownload
            // 
            this.checkBoxDownload.Location = new System.Drawing.Point(304, 88);
            this.checkBoxDownload.Name = "checkBoxDownload";
            this.checkBoxDownload.Size = new System.Drawing.Size(104, 24);
            this.checkBoxDownload.TabIndex = 7;
            this.checkBoxDownload.Text = "Download";
            this.checkBoxDownload.CheckedChanged += new System.EventHandler(this.checkBoxDownload_CheckedChanged);
            // 
            // mainMenu1
            // 
            this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem1,
            this.menuItem8,
            this.menuItem5});
            // 
            // menuItem1
            // 
            this.menuItem1.Index = 0;
            this.menuItem1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItemNewJob,
            this.menuItemLoadJob,
            this.menuItemSaveJob,
            this.menuItemSaveAs,
            this.menuItemRunBatch,
            this.menuItemViewLog,
            this.menuItemExit});
            this.menuItem1.Text = "Job";
            // 
            // menuItemNewJob
            // 
            this.menuItemNewJob.Index = 0;
            this.menuItemNewJob.Text = "New Job";
            this.menuItemNewJob.Click += new System.EventHandler(this.menuItemNewJob_Click);
            // 
            // menuItemLoadJob
            // 
            this.menuItemLoadJob.Index = 1;
            this.menuItemLoadJob.Text = "Load Job ...";
            this.menuItemLoadJob.Click += new System.EventHandler(this.menuItemLoadJob_Click);
            // 
            // menuItemSaveJob
            // 
            this.menuItemSaveJob.Index = 2;
            this.menuItemSaveJob.Text = "Save Job";
            this.menuItemSaveJob.Click += new System.EventHandler(this.menuItemSaveJob_Click);
            // 
            // menuItemSaveAs
            // 
            this.menuItemSaveAs.Index = 3;
            this.menuItemSaveAs.Text = "Save As ...";
            this.menuItemSaveAs.Click += new System.EventHandler(this.menuItemSaveAs_Click);
            // 
            // menuItemRunBatch
            // 
            this.menuItemRunBatch.Enabled = false;
            this.menuItemRunBatch.Index = 4;
            this.menuItemRunBatch.Text = "Run Batch ...";
            this.menuItemRunBatch.Visible = false;
            // 
            // menuItemViewLog
            // 
            this.menuItemViewLog.Enabled = false;
            this.menuItemViewLog.Index = 5;
            this.menuItemViewLog.Text = "View Log ...";
            this.menuItemViewLog.Click += new System.EventHandler(this.menuItemViewLog_Click);
            // 
            // menuItemExit
            // 
            this.menuItemExit.Index = 6;
            this.menuItemExit.Text = "Exit";
            this.menuItemExit.Click += new System.EventHandler(this.menuItemExit_Click);
            // 
            // menuItem8
            // 
            this.menuItem8.Index = 1;
            this.menuItem8.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItemSelectService,
            this.menuItemSetDB});
            this.menuItem8.Text = "Config";
            // 
            // menuItemSelectService
            // 
            this.menuItemSelectService.Index = 0;
            this.menuItemSelectService.Text = "Select SQL Server ...";
            this.menuItemSelectService.Click += new System.EventHandler(this.menuItemSelectService_Click);
            // 
            // menuItemSetDB
            // 
            this.menuItemSetDB.Index = 1;
            this.menuItemSetDB.Text = "Set Database Name ...";
            this.menuItemSetDB.Click += new System.EventHandler(this.menuItemSetDB_Click);
            // 
            // menuItem5
            // 
            this.menuItem5.Index = 2;
            this.menuItem5.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItemAbout});
            this.menuItem5.Text = "Help";
            // 
            // menuItemAbout
            // 
            this.menuItemAbout.Index = 0;
            this.menuItemAbout.Text = "About CXP Archive Transcoder ...";
            this.menuItemAbout.Click += new System.EventHandler(this.menuItemAbout_Click);
            // 
            // groupBoxSegments
            // 
            this.groupBoxSegments.Controls.Add(this.buttonCopySegment);
            this.groupBoxSegments.Controls.Add(this.buttonCreateSegment);
            this.groupBoxSegments.Controls.Add(this.buttonEditSegment);
            this.groupBoxSegments.Controls.Add(this.buttonDeleteSegment);
            this.groupBoxSegments.Location = new System.Drawing.Point(16, 56);
            this.groupBoxSegments.Name = "groupBoxSegments";
            this.groupBoxSegments.Size = new System.Drawing.Size(432, 144);
            this.groupBoxSegments.TabIndex = 8;
            this.groupBoxSegments.TabStop = false;
            this.groupBoxSegments.Text = "Job Segments";
            // 
            // buttonCopySegment
            // 
            this.buttonCopySegment.Enabled = false;
            this.buttonCopySegment.Location = new System.Drawing.Point(136, 20);
            this.buttonCopySegment.Name = "buttonCopySegment";
            this.buttonCopySegment.Size = new System.Drawing.Size(56, 24);
            this.buttonCopySegment.TabIndex = 7;
            this.buttonCopySegment.Text = "Copy";
            this.buttonCopySegment.Click += new System.EventHandler(this.buttonCopySegment_Click);
            // 
            // groupBoxTarget
            // 
            this.groupBoxTarget.Controls.Add(this.buttonStreamSettings);
            this.groupBoxTarget.Controls.Add(this.checkBoxDownload);
            this.groupBoxTarget.Controls.Add(this.labelOutputDir);
            this.groupBoxTarget.Controls.Add(this.textBoxOutputDir);
            this.groupBoxTarget.Controls.Add(this.textBoxBaseName);
            this.groupBoxTarget.Controls.Add(this.labelBaseName);
            this.groupBoxTarget.Controls.Add(this.checkBoxStream);
            this.groupBoxTarget.Controls.Add(this.buttonBrowse);
            this.groupBoxTarget.Controls.Add(this.checkBox1);
            this.groupBoxTarget.Location = new System.Drawing.Point(16, 273);
            this.groupBoxTarget.Name = "groupBoxTarget";
            this.groupBoxTarget.Size = new System.Drawing.Size(432, 127);
            this.groupBoxTarget.TabIndex = 9;
            this.groupBoxTarget.TabStop = false;
            this.groupBoxTarget.Text = "Output Target";
            // 
            // buttonStreamSettings
            // 
            this.buttonStreamSettings.Enabled = false;
            this.buttonStreamSettings.Location = new System.Drawing.Point(104, 89);
            this.buttonStreamSettings.Name = "buttonStreamSettings";
            this.buttonStreamSettings.Size = new System.Drawing.Size(112, 23);
            this.buttonStreamSettings.TabIndex = 8;
            this.buttonStreamSettings.Text = "Stream Settings ...";
            this.buttonStreamSettings.Click += new System.EventHandler(this.buttonStreamSettings_Click);
            // 
            // labelOutputDir
            // 
            this.labelOutputDir.Location = new System.Drawing.Point(4, 24);
            this.labelOutputDir.Name = "labelOutputDir";
            this.labelOutputDir.Size = new System.Drawing.Size(96, 16);
            this.labelOutputDir.TabIndex = 2;
            this.labelOutputDir.Text = "Output Directory:";
            this.labelOutputDir.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // textBoxOutputDir
            // 
            this.textBoxOutputDir.Location = new System.Drawing.Point(104, 24);
            this.textBoxOutputDir.Name = "textBoxOutputDir";
            this.textBoxOutputDir.Size = new System.Drawing.Size(256, 20);
            this.textBoxOutputDir.TabIndex = 1;
            this.textBoxOutputDir.TextChanged += new System.EventHandler(this.textBoxOutputDir_TextChanged);
            // 
            // textBoxBaseName
            // 
            this.textBoxBaseName.Location = new System.Drawing.Point(104, 56);
            this.textBoxBaseName.Name = "textBoxBaseName";
            this.textBoxBaseName.Size = new System.Drawing.Size(104, 20);
            this.textBoxBaseName.TabIndex = 1;
            this.textBoxBaseName.TextChanged += new System.EventHandler(this.textBoxBaseName_TextChanged);
            // 
            // labelBaseName
            // 
            this.labelBaseName.Location = new System.Drawing.Point(36, 56);
            this.labelBaseName.Name = "labelBaseName";
            this.labelBaseName.Size = new System.Drawing.Size(64, 16);
            this.labelBaseName.TabIndex = 2;
            this.labelBaseName.Text = "Base name:";
            this.labelBaseName.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // buttonBrowse
            // 
            this.buttonBrowse.Location = new System.Drawing.Point(368, 24);
            this.buttonBrowse.Name = "buttonBrowse";
            this.buttonBrowse.Size = new System.Drawing.Size(56, 24);
            this.buttonBrowse.TabIndex = 6;
            this.buttonBrowse.Text = "Browse";
            this.buttonBrowse.Click += new System.EventHandler(this.buttonBrowse_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.Enabled = false;
            this.checkBox1.Location = new System.Drawing.Point(40, 88);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(64, 24);
            this.checkBox1.TabIndex = 7;
            this.checkBox1.Text = "Stream";
            // 
            // groupBoxProfile
            // 
            this.groupBoxProfile.Controls.Add(this.comboBoxProfile);
            this.groupBoxProfile.Location = new System.Drawing.Point(16, 208);
            this.groupBoxProfile.Name = "groupBoxProfile";
            this.groupBoxProfile.Size = new System.Drawing.Size(432, 56);
            this.groupBoxProfile.TabIndex = 10;
            this.groupBoxProfile.TabStop = false;
            this.groupBoxProfile.Text = "Target Windows Media Profile";
            // 
            // comboBoxProfile
            // 
            this.comboBoxProfile.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxProfile.Location = new System.Drawing.Point(8, 21);
            this.comboBoxProfile.Name = "comboBoxProfile";
            this.comboBoxProfile.Size = new System.Drawing.Size(416, 21);
            this.comboBoxProfile.TabIndex = 0;
            this.comboBoxProfile.SelectedIndexChanged += new System.EventHandler(this.comboBoxProfile_SelectedIndexChanged);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "XML files|*.xml|All files|*.*";
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.DefaultExt = "xml";
            this.saveFileDialog1.Filter = "XML files|*.xml";
            // 
            // labelStatus
            // 
            this.labelStatus.Location = new System.Drawing.Point(16, 416);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(336, 24);
            this.labelStatus.TabIndex = 11;
            this.labelStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(448, 40);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(16, 16);
            this.button1.TabIndex = 12;
            this.button1.Text = "test";
            this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.button1_Click_2);
            // 
            // ArchiveTranscoderForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(470, 455);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.labelStatus);
            this.Controls.Add(this.listBoxSegments);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxJobTitle);
            this.Controls.Add(this.buttonStartStop);
            this.Controls.Add(this.groupBoxSegments);
            this.Controls.Add(this.groupBoxTarget);
            this.Controls.Add(this.groupBoxProfile);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Menu = this.mainMenu1;
            this.Name = "ArchiveTranscoderForm";
            this.Text = "CXP Archive Transcoder";
            this.Closing += new System.ComponentModel.CancelEventHandler(this.ArchiveTranscoderForm_Closing);
            this.groupBoxSegments.ResumeLayout(false);
            this.groupBoxTarget.ResumeLayout(false);
            this.groupBoxTarget.PerformLayout();
            this.groupBoxProfile.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		#region Transcoder Event Handlers
		private void OnBatchCompleted()
		{
			if (closing)
				return;

			if (this.InvokeRequired)
			{
				this.Invoke(new BatchCompletedDelegate(BatchCompleted));
			}
			else
			{
				BatchCompleted();
			}
		}

		private delegate void BatchCompletedDelegate();
		private void BatchCompleted()
		{
			this.buttonStartStop.Text = "Start Encoding";
			EncodingControlsEnable(true);
			this.logFile = transcoder.BatchLogFile;
			if (logFile != null)
			{
				this.menuItemViewLog.Enabled = true;
			}

			encoding = false;

			if (transcoder.BatchErrorLevel >= 7)
			{
				MessageBox.Show("Job failed: See log (available in Job menu) for details.", "Error",
					MessageBoxButtons.OK,MessageBoxIcon.Error);
				this.labelStatus.Text = "Job failed.";
			}
			else if (transcoder.BatchErrorLevel >= 5)
			{
				MessageBox.Show("There were problems found with this job: See log (available in Job menu) for details.", "Warning",
					MessageBoxButtons.OK,MessageBoxIcon.Warning);
				this.labelStatus.Text = "Job completed with warnings.";			
			}
			else
			{
				this.labelStatus.Text = "Job completed.";
			}
		}

		private void OnStatusReport(String message)
		{
			if (closing)
				return;

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

		
		private delegate void StopCompletedDelegate();
		private void StopCompletedHandler()
		{
			if (closing)
				return;

			if (this.InvokeRequired)
			{
				try
				{
					this.Invoke(new StopCompletedDelegate(StopCompleted));
				}
				catch (Exception e)
				{
					Debug.WriteLine("MainForm.StopCompletedHandler couldn't invoke: " + e.ToString());
				}
			}
			else
			{
				StopCompleted();
			}
		}

		private void StopCompleted()
		{
			if (exitAfterStop) //user exits while encoding
			{
				if (shouldPromptForSave())
				{
					if (promptForSave())
					{
						if (!save())
							return;
					}
				}
				Application.Exit();
			}
			else
			{
				this.labelStatus.Text = "Job Interrupted.";
				this.Cursor = Cursors.Default;
				this.buttonStartStop.Text = "Start Encoding";
				this.buttonStartStop.Enabled = true;
				EncodingControlsEnable(true);
				encoding = false;
			}
		}


		#endregion Transcoder Event Handlers

		#region Windows Form Event Handlers

		#region Form Event Handlers

		/// <summary>
		/// If the user clicks the X in the upper right of the form, we get this chance to shut down and save.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void ArchiveTranscoderForm_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			//Flag threads that would normally invoke to the form thread not to do so:
			closing=true;
			
			if ((transcoder != null) && (this.encoding))
			{
				this.Cursor = Cursors.WaitCursor;
				transcoder.Stop();
			}		

			if (shouldPromptForSave())
			{
				if (promptForSave())
				{
					save();
				}
			}
		}

		#endregion Form Event Handlers

		#region Menu Event Handlers

		private void menuItemExit_Click(object sender, System.EventArgs e)
		{
			if (encoding)
			{
				this.Cursor = Cursors.WaitCursor;
				transcoder.AsyncStop();
				this.buttonStartStop.Enabled = false;
				this.buttonStartStop.Text = "Stopping ...";
				exitAfterStop = true;
			}
			else
			{
				if (shouldPromptForSave())
				{
					if (promptForSave())
					{
						if (!save())
							return;
					}
				}
				Application.Exit();
			}
		}

						
		private void menuItemLoadJob_Click(object sender, System.EventArgs e)
		{
			if (shouldPromptForSave())
			{
				if (promptForSave())
				{
					if (!save())
						return;
				}
			}

			if (jobDir != "")
			{
				openFileDialog1.InitialDirectory = jobDir;
			}
			if (DialogResult.OK == this.openFileDialog1.ShowDialog(this))
			{
				jobDir = this.openFileDialog1.FileName;
				loadJob(this.openFileDialog1.FileName);
			}
		}

		private void menuItemSaveJob_Click(object sender, System.EventArgs e)
		{
			save();
		}


		private void menuItemSaveAs_Click(object sender, System.EventArgs e)
		{
			saveAs();
		}


		private void menuItemAbout_Click(object sender, System.EventArgs e)
		{
			frmAbout about = new frmAbout();
			about.ShowDialog(this);
		}

		private void menuItemSelectService_Click(object sender, System.EventArgs e)
		{
			SetServiceForm setForm = new SetServiceForm();
			setForm.textBoxHostName.Text = transcoder.SQLHost;
			if (DialogResult.OK == setForm.ShowDialog(this))
			{
				if (transcoder.SQLHost != setForm.textBoxHostName.Text.Trim())
				{
					confDataCache.Stop();
					transcoder.SQLHost = setForm.textBoxHostName.Text.Trim();
					confDataCache.SqlServer = transcoder.SQLHost;
					confDataCache.Lookup();
				}
			}
		}

		private void menuItemSetDB_Click(object sender, System.EventArgs e)
		{
			SetDBNameForm dbNameForm = new SetDBNameForm();
			dbNameForm.DBName = transcoder.DBName;
			if (DialogResult.OK == dbNameForm.ShowDialog(this))
			{
				if (transcoder.DBName != dbNameForm.DBName)
				{
					confDataCache.Stop();
					transcoder.DBName = dbNameForm.DBName;
					confDataCache.DbName = transcoder.DBName;
					confDataCache.Lookup();
				}
			}
		}

		private void menuItemViewLog_Click(object sender, System.EventArgs e)
		{
			if (logFile==null)
			{
				MessageBox.Show("No log file is set.");
				return;
			}
			if (!File.Exists(logFile))
			{
				MessageBox.Show("Log file does not exist: " + logFile);
				return;
			}
			Process proc = new Process();
			proc.StartInfo.FileName = logFile;
			proc.StartInfo.Verb = "Open";
			proc.Start();
			
		}

		private void menuItemNewJob_Click(object sender, System.EventArgs e)
		{
			if (shouldPromptForSave())
			{
				if (promptForSave())
				{
					if (!save())
						return;
				}
			}
			this.textBoxJobTitle.Text = "";
			this.textBoxBaseName.Text = "";
			this.listBoxSegments.Items.Clear();
			this.checkBoxStream.Checked = false;
			this.checkBoxDownload.Checked = false;
			this.currentDoc = null;
			this.buttonCopySegment.Enabled = false;
			this.buttonEditSegment.Enabled = false;
			this.buttonDeleteSegment.Enabled = false;

			setFormTitle();
		}

		#endregion Menu Event Handlers

		#region Button Event Handlers

		private void buttonCreateSegment_Click(object sender, System.EventArgs e)
		{
			if (transcoder.SQLHost == "")
			{
				this.menuItemSelectService_Click(sender, e);
			}

			SegmentForm sform = new SegmentForm(confDataCache);

			if (!confDataCache.Working.WaitOne(0,false))
			{
				StreamLookupProgressForm slpform = new StreamLookupProgressForm(confDataCache);
				DialogResult dr = slpform.ShowDialog(this);
				if (dr == DialogResult.Cancel)
				{
					return;
				}
			}

			if ((confDataCache.Done==true) && (confDataCache.SqlConnectError == false))
			{
				sform.Initialize(); 
				DialogResult result = sform.ShowDialog(this);
				if (result == DialogResult.OK)
				{
					SegmentWrapper wrapper = new SegmentWrapper(sform.Segment);
					AddSortedSegment(wrapper);
					dirtyBit = true;
				}
			}
			else
			{
				MessageBox.Show("SQL Server connection failed.  Please set or verify SQL Server host and database name using the 'Config' menu.");
				return;
			}
		}

		private void buttonDeleteSegment_Click(object sender, System.EventArgs e)
		{
			this.listBoxSegments.Items.RemoveAt(listBoxSegments.SelectedIndex);
			buttonDeleteSegment.Enabled = false;
			buttonEditSegment.Enabled = false;
			buttonCopySegment.Enabled = false;
			dirtyBit = true;
		}

		private void buttonEditSegment_Click(object sender, System.EventArgs e)
		{
			SegmentWrapper wrapper = (SegmentWrapper)listBoxSegments.SelectedItem;
			SegmentForm sform = new SegmentForm(confDataCache);

			if (!confDataCache.Working.WaitOne(0,false))
			{
				StreamLookupProgressForm slpform = new StreamLookupProgressForm(confDataCache);
				DialogResult dr = slpform.ShowDialog(this);
				if (dr == DialogResult.Cancel)
				{
					return;
				}
			}

			if ((confDataCache.Done==true) && (confDataCache.SqlConnectError == false))
			{
				sform.Initialize();
				bool success = sform.LoadSegment(wrapper.Segment);

				if (!success)
				{
					MessageBox.Show("Warning: This segment does not appear to match a conference in the database.", "Conference missing", 
						MessageBoxButtons.OK,MessageBoxIcon.Warning);
				}

				DialogResult result = sform.ShowDialog(this);
				if (result == DialogResult.OK)
				{
					listBoxSegments.Items.Remove(wrapper);
					wrapper = new SegmentWrapper(sform.Segment);
					AddSortedSegment(wrapper);
					buttonDeleteSegment.Enabled = false;
					buttonEditSegment.Enabled = false;
					buttonCopySegment.Enabled = false;
					dirtyBit = true;
				}
			}
			else
			{
				MessageBox.Show("SQL Server connection failed.  Please set or verify SQL Server host and database name using the 'Config' menu.");
				return;			
			}
		}


		/// <summary>
		/// Validate inputs and if things look good, construct the batch class and start encoding.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void buttonStartStop_Click(object sender, System.EventArgs e)
		{
			if (encoding)
			{
				this.Cursor = Cursors.WaitCursor;
				transcoder.AsyncStop();
				this.buttonStartStop.Enabled = false;
				this.buttonStartStop.Text = "Stopping ...";
			}
			else
			{
				this.labelStatus.Text = "";
				if (!transcoder.VerifyDBConnection())
				{
					MessageBox.Show("SQL Server connection failed.  Please set or verify SQL Server host and database name.","Error");
					return;
				}
				
				//Make sure we have all the mandatory fields
				String errorMsg;
				if (!FormDataComplete(out errorMsg))
				{
					MessageBox.Show("The input data is incomplete: \r\n" + errorMsg,"Data is incomplete", 
						MessageBoxButtons.OK,MessageBoxIcon.Error);
					return;
				}

				batch = new ArchiveTranscoderBatch();
				this.PopulateBatchFromForm();
				transcoder.LoadJobs(batch);

				//confirm file overwrites if any
				if (transcoder.WillOverwriteFiles)
				{
					DialogResult dr = MessageBox.Show("Warning!  This operation will overwrite existing files and directories.  " +
						"\r\n\r\nAre you sure you want to proceed?","Confirm Files/Directories Overwrite",
						MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
					if (dr != DialogResult.Yes)
					{
						return;
					}
				}

				//Warn if decks don't exist or are unmatched.
				String deckWarnings = transcoder.GetDeckWarnings();
				if ((deckWarnings != null) && (deckWarnings != ""))
				{
					DialogResult dr = MessageBox.Show(this, deckWarnings + 
						"\r\nThis problem may prevent Archive Transcoder from correctly building slide images. " +
						"\r\n\r\nAre you sure you want to proceed?","Missing, Unmatched or Invalid Deck",MessageBoxButtons.YesNo,
						MessageBoxIcon.Warning);
					if (dr != DialogResult.Yes)
					{
						return;
					}
				}

				String errMsg;
				errMsg=transcoder.Encode();
				if (errMsg != null)
				{
					MessageBox.Show(errMsg,"Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
					return;
				}
				this.buttonStartStop.Text = "Stop Encoding";
				encoding = true;

				EncodingControlsEnable(false);
				this.menuItemViewLog.Enabled = false;

			}
		}

		private void buttonBrowse_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.OK == folderBrowserDialog1.ShowDialog())
			{
				this.textBoxOutputDir.Text = folderBrowserDialog1.SelectedPath;
			}
		}

		private void buttonCopySegment_Click(object sender, System.EventArgs e)
		{
			//Copy the selected segment, setting the start time to the end time of the segment to be copied.
			SegmentWrapper wrapper = (SegmentWrapper)listBoxSegments.SelectedItem;
			SegmentWrapper newWrapper = new SegmentWrapper(GetNextSegment(wrapper.Segment));
			AddSortedSegment(newWrapper);
			buttonDeleteSegment.Enabled = false;
			buttonEditSegment.Enabled = false;
			buttonCopySegment.Enabled = false;
			dirtyBit = true;
		}

		private void buttonStreamSettings_Click(object sender, System.EventArgs e)
		{
			
			StreamSettingsForm streamSettingsForm = new StreamSettingsForm();
			if (streamTarget != null)
			{
				streamSettingsForm.SetJobTarget(streamTarget);
			}
			if (DialogResult.OK == streamSettingsForm.ShowDialog(this))
			{
				streamTarget = streamSettingsForm.GetJobTarget();
				dirtyBit = true;
			}
		}


		//button for testing new code
		private void button1_Click_2(object sender, System.EventArgs e)
		{
        }

		#endregion Button Event Handlers

		#region ListBox Event Handlers
		private void listBoxSegments_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.buttonEditSegment.Enabled = true;
			this.buttonDeleteSegment.Enabled = true;
			this.buttonCopySegment.Enabled = true;
		}

		/// <summary>
		/// Double click on a segment is the same clicking edit
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void listBoxSegments_DoubleClick(object sender, System.EventArgs e)
		{
			if (buttonEditSegment.Enabled)
				this.buttonEditSegment_Click(sender,e);
		}

		#endregion ListBox Event Handlers

		#region CheckBox Event Handlers
		private void checkBoxStream_CheckedChanged(object sender, System.EventArgs e)
		{
			this.buttonStreamSettings.Enabled = this.checkBoxStream.Checked;
			dirtyBit = true;
		}

		private void checkBoxDownload_CheckedChanged(object sender, System.EventArgs e)
		{
			dirtyBit = true;
		}


		#endregion CheckBox Event Handlers

		#region ComboBox Event Handlers

		private void comboBoxProfile_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			dirtyBit = true;
		}

		#endregion ComboBox Event Handlers

		#region TextBox Event Handlers

		private void textBoxJobTitle_TextChanged(object sender, System.EventArgs e)
		{
			dirtyBit = true;
		}

		private void textBoxOutputDir_TextChanged(object sender, System.EventArgs e)
		{
			dirtyBit = true;
		}

		private void textBoxBaseName_TextChanged(object sender, System.EventArgs e)
		{
			dirtyBit = true;
		}


		#endregion TextBox Event Handlers

		#endregion Windows Form Event Handlers

		#region Private Methods

		/// <summary>
		/// Check the WMP version and warn if it isn't 10 or larger.
		/// </summary>
		private void checkWmpVersion()
		{
			//If the user has set the bit to tell us never to warn again, do nothing.
			if (shouldShowWmpWarnings())
			{
				// I have a feeling this isn't very robust ...
				// look for HKLM\Software\Microsoft\MediaPlayer\10.0
				RegistryKey wmpKey = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\MediaPlayer");
				if (wmpKey != null)
				{
					RegistryKey k = wmpKey.OpenSubKey("10.0");
					if (k == null)
					{
						k = wmpKey.OpenSubKey("11.0");
						if (k == null)
						{
							k = wmpKey.OpenSubKey("12.0");
							if (k == null)
							{
								showWmpWarning();
							}
						}
					}
				}
				else
				{
					showWmpWarning();
				}
			}
		}

		/// <summary>
		/// Check the registry entry to see if the user has chosen to never see the WMP warnings dialog.
		/// </summary>
		/// <returns></returns>
		private bool shouldShowWmpWarnings()
		{
			try
			{
				RegistryKey BaseKey = Registry.CurrentUser.OpenSubKey(Constants.AppRegKey, true);
				if ( BaseKey == null) 
				{ //no app configuration yet.. first run.
					return true;
				}
				else
				{
					return Convert.ToBoolean(BaseKey.GetValue("ShowWmpWarning","true"));
				}
			}
			catch
			{
				return true;
			}
		}

		/// <summary>
		/// Show a dialog that warns the user about the need for WMP 10 or later.  The dialog has a checkbox
		/// that the user can tick to never show this warning again.  If the tick is entered, write a reg entry.
		/// </summary>
		private void showWmpWarning()
		{
			WmpWarningForm wwf = new WmpWarningForm();
			if (DialogResult.OK == wwf.ShowDialog(this))
			{
				if (wwf.DisableWarning)
				{
					try
					{
						RegistryKey BaseKey = Registry.CurrentUser.OpenSubKey(Constants.AppRegKey, true);
						if ( BaseKey == null) 
						{
							BaseKey = Registry.CurrentUser.CreateSubKey(Constants.AppRegKey);
						}
						BaseKey.SetValue("ShowWmpWarning",false);
					}
					catch
					{
						Debug.WriteLine("Exception while saving configuration.");
					}
				}
			}
		}

		private void setToolTips()
		{
			ToolTip tt = new ToolTip();
			tt.SetToolTip(this.checkBoxDownload,"Package files for stand-alone or offline use");
			tt.SetToolTip(this.checkBoxStream,"Prepare files for server-based streaming scenarios");
			tt.SetToolTip(this.listBoxSegments,"A segment consists of one continuous time period, one video source, one or more audio sources and an optional presentation source." +
				"\r\nMultiple segments defined for a job will be combined to make a continuous Windows Media output.");
			tt.SetToolTip(this.textBoxBaseName,"A short string to be used as the base part of file names in the output");
			tt.SetToolTip(this.labelBaseName,"A short string to be used as the base part of file names in the output");
			tt.SetToolTip(this.textBoxJobTitle,"A descriptive name for the archive");
			tt.SetToolTip(this.textBoxOutputDir,"Where to put the transcoder output");
			tt.SetToolTip(this.labelOutputDir,"Where to put the transcoder output");
		}

		/// <summary>
		/// Populate the combobox from which the user chooses a Windows Media Profile.
		/// The first item is always "No Recompression".  The next few items will be some standard
		/// system profiles.  These will be followed by any custom profiles which the user has specified.
		/// </summary>
		private void FillWMProfileList()
		{
			this.comboBoxProfile.Items.Add(new ProfileWrapper(ProfileWrapper.EntryType.NoRecompression,"No Recompression","norecompression"));
			//PRI3: It would be better to use the guids instead of the numerals to specify system profiles.
			this.comboBoxProfile.Items.Add(new ProfileWrapper(ProfileWrapper.EntryType.System,"256 Kbps System Profile","9"));
			this.comboBoxProfile.Items.Add(new ProfileWrapper(ProfileWrapper.EntryType.System,"384 Kbps System Profile","10"));
			this.comboBoxProfile.Items.Add(new ProfileWrapper(ProfileWrapper.EntryType.System,"768 Kbps System Profile","11"));
			this.comboBoxProfile.SelectedIndex=0;

			// Add the custom profiles found in prx files in the application directory:
			string[] prxFiles = Directory.GetFileSystemEntries(Path.GetDirectoryName(Application.ExecutablePath), "*.prx");
			DataSet ds = new DataSet();
			string pname;
			for (int i =0;i<prxFiles.Length;i++)
			{	
				try 
				{
					ds.Clear();
					ds.ReadXml(prxFiles[i]);
					// By ReadXml's inference rules, the following should get
					// the profile name from xml such as:
					//   <profile ... name="Profile Name" ... > ... </profile>
					pname = (string) ds.Tables["profile"].Rows[0]["name"];
				} 
				catch 
				{
					MessageBox.Show("Failed to parse Windows Media Profile: " + prxFiles[i] + ".  " +
						"You may want to repair this file, or remove it from the application directory.  " +
						"This profile will be ignored. " , "Invalid Profile");
					continue;
				}
				comboBoxProfile.Items.Add(new ProfileWrapper(ProfileWrapper.EntryType.Custom,pname  + " (Custom Profile)",prxFiles[i]));
			}
		}


		private void loadJob(String filename)
		{
			batch = Utility.ReadBatchFile(filename);
			if (batch == null)
			{
				MessageBox.Show("Error: Failed to read and parse batch file: " + filename, "Error", MessageBoxButtons.OK,
					MessageBoxIcon.Error);
				return;
			}
			if (batch.Job.Length == 0)
			{
				MessageBox.Show("Error: Invalid batch file.  (It contains zero jobs.)", "Error", MessageBoxButtons.OK,
					MessageBoxIcon.Error);
				batch = null;
				streamTarget = null;
				return;
			}
			if (batch.Job.Length > 1)
			{
				MessageBox.Show("Note: The selected batch file contains multiple jobs.  \n\r" +
					"There is no UI support for modifying batch files with multiple jobs. \n\r" +
					"Make any modifications in the XML with a text editor, and use the 'Run Batch' command to execute the job.", 
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				batch = null;
				streamTarget = null;
				return;
			}

			try
			{
				this.textBoxJobTitle.Text = batch.Job[0].ArchiveName;
				this.textBoxOutputDir.Text = batch.Job[0].Path;
				this.textBoxBaseName.Text = batch.Job[0].BaseName;
				this.listBoxSegments.Items.Clear();
				this.buttonEditSegment.Enabled = false;
				this.buttonCopySegment.Enabled = false;
				this.buttonDeleteSegment.Enabled = false;

				ArrayList segmentList = new ArrayList();
				foreach (ArchiveTranscoderJobSegment s in batch.Job[0].Segment)
				{
					segmentList.Add(new SegmentWrapper(s));
				}
				segmentList.Sort();
				this.listBoxSegments.Items.AddRange((object[])segmentList.ToArray(typeof(object)));

				this.comboBoxProfile.SelectedIndex=0;
				if (batch.Job[0].WMProfile != null)
				{
					bool foundProfile = false;
                    string profile = batch.Job[0].WMProfile;
                    uint junk;
                    if (!uint.TryParse(profile,out junk)) {
                        if ((profile.ToLower() != "norecompression") && (!File.Exists(profile))) {
                            //If the prx file isn't found try looking in the app directory.
                            //This is to hack around an issue with working between x64 and x86 systems with different Program Files paths.
                            System.Reflection.Assembly a = System.Reflection.Assembly.GetExecutingAssembly();
                            String appDir = System.IO.Path.GetDirectoryName(a.Location);
                            profile = Path.Combine(appDir, Path.GetFileName(profile));
                        }
                    }

					foreach(Object o in this.comboBoxProfile.Items)
					{
						ProfileWrapper pw = (ProfileWrapper)o;

						if (pw.ProfileValue==profile)
						{
							foundProfile = true;
							this.comboBoxProfile.SelectedItem = o;
							break;
						}
					}
					//PRI3: we could be slightly smarter about this by also attempting to make a match by name.
					if (!foundProfile)
					{
						MessageBox.Show("The Windows Media Profile specified in the job was not found: \r\n   " + profile +
							"\r\n\r\nPlease verify the profile setting.", "Profile Not Found", 
							MessageBoxButtons.OK, MessageBoxIcon.Warning);
					}
				}
				this.checkBoxDownload.Checked = this.checkBoxStream.Checked = false;
				streamTarget = null;
				if (batch.Job[0].Target != null)
				{
					foreach (ArchiveTranscoderJobTarget t in batch.Job[0].Target)
					{
						if (t.Type.ToLower().Trim() == "download")
						{
							this.checkBoxDownload.Checked = true;
						}
						if (t.Type.ToLower().Trim() == "stream")
						{
							this.checkBoxStream.Checked = true;
							streamTarget = t;
						}
					}
				}

				//use the server if specified, else keep the current server setting.
				String oldSqlHost = transcoder.SQLHost;
				String oldDBName = transcoder.DBName;
				String newSqlHost = oldSqlHost;
				String newDBName = oldDBName;
				if ((batch.Server != null) && (batch.Server != ""))
				{
					newSqlHost = batch.Server;
				}
				if ((batch.DatabaseName != null) && (batch.DatabaseName != ""))
				{
					newDBName = batch.DatabaseName;
				}

				//If either sqlHost or DBName changed, refresh the conference data, but not if we are just launching, 
				// in which case it will be invoked shortly.
				if ((newSqlHost != oldSqlHost) || (newDBName != oldDBName))
				{
					if (confDataCache !=null)
					{
						confDataCache.Stop();
					}
					
					transcoder.SQLHost = newSqlHost;
					transcoder.DBName = newDBName;

					if (confDataCache != null)
					{
						confDataCache.SqlServer = transcoder.SQLHost;
						confDataCache.DbName = transcoder.DBName;
						confDataCache.Lookup();
					}
				}

				this.currentDoc = filename;
				this.setFormTitle();
				dirtyBit = false;
			}
			catch
			{
				MessageBox.Show("The file format is incorrect (possibly an old format).  " +
					"Verify job parameters and resave.", "Warning", MessageBoxButtons.OK,
					MessageBoxIcon.Warning);
			}

		}

		private bool save()
		{
			bool saved = false;
			String errorMsg;
			if (FormDataComplete(out errorMsg))
			{
				if (this.currentDoc != null)
				{
					PopulateBatchFromForm();
					Utility.WriteBatch(batch,currentDoc);
					saved = true;
					jobDir = currentDoc;
					dirtyBit = false;
				}
				else
				{
					saved = saveAs();
				}
			}
			else
			{
				MessageBox.Show("Before saving, be sure all form data is complete: \r\n" + errorMsg, "Data is incomplete.", 
					MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}		
			return saved;
		}

		/// <summary>
		/// Do Save As, returning true if the job was saved.
		/// </summary>
		/// <returns></returns>
		private bool saveAs()
		{
			bool saved = false;
			String errorMsg;
			if (FormDataComplete(out errorMsg))
			{
				if (this.jobDir != "")
				{
					saveFileDialog1.InitialDirectory = jobDir;
				}
				if (DialogResult.OK == saveFileDialog1.ShowDialog(this))
				{
					PopulateBatchFromForm();
					Utility.WriteBatch(batch,saveFileDialog1.FileName);
					saved = true;
					jobDir = saveFileDialog1.FileName;
					currentDoc = saveFileDialog1.FileName;
					dirtyBit = false;
					setFormTitle();
				}
			}
			else
			{
				MessageBox.Show("Before saving, be sure all form data is complete: \r\n" + errorMsg, "Data is incomplete",
					MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}	
			return saved;
		}

		private bool FormDataComplete(out String errorMsg)
		{
			StringBuilder error = new StringBuilder();
			if (this.textBoxJobTitle.Text.Trim() == "")
			{
				error.Append("  Job Title must be specified. \r\n");
			}
			if (this.textBoxBaseName.Text.Trim() == "")
			{
				error.Append("  Job Base Name must be specified. \r\n");	
			}
			if (this.textBoxOutputDir.Text.Trim() == "")
			{				
				error.Append("  Job Output Directory must be specified. \r\n");
			}
			if (this.listBoxSegments.Items.Count == 0)
			{				
				error.Append("  At least one job segment must be specified. \r\n");
			}

			if ((!this.checkBoxStream.Checked) && (!this.checkBoxDownload.Checked))
			{
				error.Append("  At least one output target (stream/download) must be selected. \r\n");
			}

			errorMsg = error.ToString();

			if (errorMsg == "")
				return true;

			return false;
		}

		private void PopulateBatchFromForm()
		{
			batch = new ArchiveTranscoderBatch();
			batch.Job = new ArchiveTranscoderJob[1];
			batch.Job[0] = new ArchiveTranscoderJob();
			batch.Job[0].ArchiveName = this.textBoxJobTitle.Text.Trim();
			batch.Job[0].BaseName = this.textBoxBaseName.Text.Trim();
			batch.Job[0].Path = this.textBoxOutputDir.Text.Trim();
			ProfileWrapper pw = (ProfileWrapper)comboBoxProfile.SelectedItem;
			batch.Job[0].WMProfile = pw.ProfileValue;
			batch.Job[0].Segment = new ArchiveTranscoderJobSegment[this.listBoxSegments.Items.Count];
			for(int i=0;i< this.listBoxSegments.Items.Count; i++)
			{
				batch.Job[0].Segment[i] = ((SegmentWrapper)this.listBoxSegments.Items[i]).Segment;
			}

			if ((this.checkBoxStream.Checked) && (this.checkBoxDownload.Checked))
			{
				batch.Job[0].Target = new ArchiveTranscoderJobTarget[2];
				batch.Job[0].Target[0] = new ArchiveTranscoderJobTarget();
				batch.Job[0].Target[0].Type = "download";
				if (streamTarget != null)
					batch.Job[0].Target[1] = streamTarget;
				else
					batch.Job[0].Target[1] = getDefaultStreamTarget();
				batch.Job[0].Target[1].Type = "stream";
			}
			else if (this.checkBoxDownload.Checked)
			{
				batch.Job[0].Target = new ArchiveTranscoderJobTarget[1];
				batch.Job[0].Target[0] = new ArchiveTranscoderJobTarget();
				batch.Job[0].Target[0].Type = "download";
			}
			else if (this.checkBoxStream.Checked)
			{
				batch.Job[0].Target = new ArchiveTranscoderJobTarget[1];
				if (streamTarget != null)
					batch.Job[0].Target[0] = streamTarget;
				else
					batch.Job[0].Target[0] = getDefaultStreamTarget();
				batch.Job[0].Target[0].Type = "stream";
			}
			
			if (transcoder.SQLHost != "")
				batch.Server = transcoder.SQLHost;
			if (transcoder.DBName != "")
				batch.DatabaseName = transcoder.DBName;

		}

		/// <summary>
		/// Restore default streaming target parameters from the registry if they exist.
		/// </summary>
		/// <returns></returns>
		private ArchiveTranscoderJobTarget getDefaultStreamTarget()
		{
			ArchiveTranscoderJobTarget target = new ArchiveTranscoderJobTarget();
			target.Type = "stream";
			try
			{
				RegistryKey BaseKey = Registry.CurrentUser.OpenSubKey(Constants.AppRegKey, true);
				if ( BaseKey == null) 
				{ //no configuration yet.. first run.
					//Debug.WriteLine("No registry configuration found.");
				}
				else
				{
					target.SlideBaseUrl = Convert.ToString(BaseKey.GetValue("SlideBaseUrl",""));
					target.WmvUrl = Convert.ToString(BaseKey.GetValue("WmvUrl",""));
					target.PresentationUrl = Convert.ToString(BaseKey.GetValue("PresentationUrl",""));
					target.AsxUrl = Convert.ToString(BaseKey.GetValue("AsxUrl",""));
					target.CreateAsx = "False";
					target.CreateWbv = "False";
					if (Convert.ToBoolean(BaseKey.GetValue("CreateAsx","False")))
					{
						target.CreateAsx = "True";
					}
					if (Convert.ToBoolean(BaseKey.GetValue("CreateWbv","False")))
					{
						target.CreateWbv = "True";
					}
				}
			}
			catch (Exception e)
			{
				Debug.WriteLine("Exception while reading registry: " + e.ToString());
			}

			return target;
		}

		/// <summary>
		/// Enable/disable all controls when we start or end a transcoding job
		/// </summary>
		/// <param name="enable"></param>
		private void EncodingControlsEnable(bool enable)
		{
			this.textBoxJobTitle.Enabled = enable;
			this.buttonCreateSegment.Enabled = enable;
			this.buttonEditSegment.Enabled = enable;
			this.buttonCopySegment.Enabled = enable;
			this.buttonDeleteSegment.Enabled = enable;
			if (enable)
			{
				if (this.listBoxSegments.SelectedItem == null)
				{
					this.buttonEditSegment.Enabled = false;
					this.buttonCopySegment.Enabled = false;
					this.buttonDeleteSegment.Enabled = false;
				}
			}
			this.listBoxSegments.Enabled = enable;
			this.comboBoxProfile.Enabled = enable;
			this.textBoxBaseName.Enabled = enable;
			this.textBoxOutputDir.Enabled = enable;
			this.buttonBrowse.Enabled = enable;
			this.menuItemLoadJob.Enabled = enable;
			this.menuItemSelectService.Enabled = enable;
			this.menuItemNewJob.Enabled = enable;
			this.menuItemSetDB.Enabled = enable;
			this.checkBoxStream.Enabled = enable;
			this.checkBoxDownload.Enabled = enable;
			this.buttonStreamSettings.Enabled = enable;
			if ((enable) && (!this.checkBoxStream.Checked))
				this.buttonStreamSettings.Enabled = false;
			else
				this.buttonStreamSettings.Enabled = enable;

		}

		/// <summary>
		/// Add wrapper to listBoxSegments, making sure that it is in order by startTime.
		/// </summary>
		/// <param name="wrapper"></param>
		private void AddSortedSegment(SegmentWrapper wrapper)
		{
			ArrayList segList = new ArrayList();
			foreach(object obj in this.listBoxSegments.Items)
			{
				if (obj is SegmentWrapper)
				{
					segList.Add(obj);
				}
			}
			segList.Add(wrapper);
			segList.Sort();
			this.listBoxSegments.Items.Clear();
			this.listBoxSegments.Items.AddRange((object[])segList.ToArray(typeof(object)));
		}



		/// <summary>
		/// Copy segment, and set the copy's start time to the original end time, and the new end time
		/// 1 minute later.
		/// </summary>
		/// <param name="segment"></param>
		/// <returns></returns>
		private ArchiveTranscoderJobSegment GetNextSegment(ArchiveTranscoderJobSegment segment)
		{
		
			ArchiveTranscoderJobSegment newSegment = new ArchiveTranscoderJobSegment();
			if (segment.PresentationDescriptor != null)
			{
				newSegment.PresentationDescriptor = new ArchiveTranscoderJobSegmentPresentationDescriptor();
				newSegment.PresentationDescriptor.PresentationCname = segment.PresentationDescriptor.PresentationCname;
				newSegment.PresentationDescriptor.PresentationFormat = segment.PresentationDescriptor.PresentationFormat;
                if (segment.PresentationDescriptor.VideoDescriptor != null) {
                    newSegment.PresentationDescriptor.VideoDescriptor = new ArchiveTranscoderJobSegmentVideoDescriptor();
                    newSegment.PresentationDescriptor.VideoDescriptor.VideoCname = segment.PresentationDescriptor.VideoDescriptor.VideoCname;
                    newSegment.PresentationDescriptor.VideoDescriptor.VideoName = segment.PresentationDescriptor.VideoDescriptor.VideoName;
                }
			}
			newSegment.StartTime = segment.EndTime;
            if (segment.VideoDescriptor != null)
            {
                newSegment.VideoDescriptor = new ArchiveTranscoderJobSegmentVideoDescriptor();
                newSegment.VideoDescriptor.VideoCname = segment.VideoDescriptor.VideoCname;
                newSegment.VideoDescriptor.VideoName = segment.VideoDescriptor.VideoName;
            }
            if (segment.Flags != null)
            {
                newSegment.Flags = (string[])segment.Flags.Clone();
            }
            newSegment.AudioDescriptor = new ArchiveTranscoderJobSegmentAudioDescriptor[segment.AudioDescriptor.Length];
			for (int i = 0; i<segment.AudioDescriptor.Length; i++)
			{
				newSegment.AudioDescriptor[i] = new ArchiveTranscoderJobSegmentAudioDescriptor();
				newSegment.AudioDescriptor[i].AudioCname = segment.AudioDescriptor[i].AudioCname;
				newSegment.AudioDescriptor[i].AudioName = segment.AudioDescriptor[i].AudioName;
			}
			if (segment.SlideDecks != null)
			{
				newSegment.SlideDecks = new ArchiveTranscoderJobSlideDeck[segment.SlideDecks.Length];
				for (int i = 0; i<segment.SlideDecks.Length; i++)
				{
					newSegment.SlideDecks[i] = new ArchiveTranscoderJobSlideDeck();
					newSegment.SlideDecks[i].DeckGuid = segment.SlideDecks[i].DeckGuid;
					newSegment.SlideDecks[i].Path = segment.SlideDecks[i].Path;
					newSegment.SlideDecks[i].Title = segment.SlideDecks[i].Title;
				}
			}
			DateTime newEndTime = DateTime.Parse(segment.EndTime) + TimeSpan.FromSeconds(60);
			//newSegment.EndTime = newEndTime.ToString(Constants.dtformat);
            newSegment.EndTime = Utility.GetLocalizedDateTimeString(newEndTime, Constants.timeformat);
			return newSegment;
		}

		private void restoreRegSettings()
		{
			jobDir = "";
			String outDir = @"C:\";
			try
			{
				RegistryKey BaseKey = Registry.CurrentUser.OpenSubKey(Constants.AppRegKey, true);
				if ( BaseKey == null) 
				{ //no configuration yet.. first run.
					//Debug.WriteLine("No registry configuration found.");
					return;
				}
				jobDir = Convert.ToString(BaseKey.GetValue("JobDir",jobDir));
				outDir = Convert.ToString(BaseKey.GetValue("OutDir",outDir));
				this.textBoxOutputDir.Text = outDir;
			}
			catch (Exception e)
			{
				Debug.WriteLine("Exception while reading registry: " + e.ToString());
			}
		}

		private void saveRegSettings()
		{
			String outDir = this.textBoxOutputDir.Text.Trim();
			try
			{
				RegistryKey BaseKey = Registry.CurrentUser.OpenSubKey(Constants.AppRegKey, true);
				if ( BaseKey == null) 
				{
					BaseKey = Registry.CurrentUser.CreateSubKey(Constants.AppRegKey);
				}
				BaseKey.SetValue("JobDir",jobDir);
				BaseKey.SetValue("OutDir",outDir);
			}
			catch
			{
				Debug.WriteLine("Exception while saving configuration.");
			}
		}


		/// <summary>
		/// Look in the currently defined temp location for files that appear to belong to ArchiveTranscoder.
		/// Offer to delete them if found.  Note: there should never be any stray files left if things are
		/// working correctly.
		/// </summary>
		private void cleanUpOldTempFiles()
		{
			String filelist = Utility.GetExistingTempFilesAndDirs();
			if ((filelist != null) && (filelist != ""))
			{
				DialogResult dr = MessageBox.Show(this,"The following temporary items appear to be left over from a previous run of ArchiveTranscoder: \r\n" +
					filelist + "\r\nWould you like to delete these files/directories now?","Delete Temporary Items?", 
					MessageBoxButtons.YesNo, MessageBoxIcon.Question);
				if (dr == DialogResult.Yes)
				{
					Utility.DeleteExistingTempFilesAndDirs();
				}
			}
		}


		/// <summary>
		/// Called before new job or load job.  Examine current state of the form, and return
		/// true if it is such that we should ask the user to save before proceeding.
		/// </summary>
		/// <returns></returns>
		/// We will prompt for save if form data is complete enough, and if any of the data has changed.
		private bool shouldPromptForSave()
		{
			String junk;
			if (this.FormDataComplete(out junk))
			{
				if (dirtyBit)
					return true;
			}
			return false;
		}

		private bool promptForSave()
		{
			DialogResult dr = MessageBox.Show(this, "Do you want to save the current job settings?", "Do you want to save?",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			return (dr == DialogResult.Yes);
		}

		private void setFormTitle()
		{
			if (currentDoc != null)
			{
				String filename = Path.GetFileName(currentDoc);
				this.Text = "CXP Archive Transcoder: " + filename;
			}
			else
			{
				this.Text = "CXP Archive Transcoder: New Job";
			}
		}
		#endregion Private Methods

	}

}
