using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

namespace ArchiveTranscoder
{
	/// <summary>
	/// AV Preview UI to help the user find start/end times
	/// </summary>
	public class PreviewForm : System.Windows.Forms.Form
	{
		#region Windows Forms Members

		private System.Windows.Forms.Label labelVideoSource;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.CheckedListBox checkedListBoxAudioSources;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.GroupBox groupBoxStartDuration;
		private System.Windows.Forms.TextBox textBoxStartTime;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox textBoxDuration;
		private System.Windows.Forms.Button buttonBuild;
		private System.Windows.Forms.Button buttonMarkOut;
		private System.Windows.Forms.Button buttonMarkIn;
		private System.Windows.Forms.TextBox textBoxMarkIn;
		private System.Windows.Forms.TextBox textBoxMarkOut;
		private AxWMPLib.AxWindowsMediaPlayer axWindowsMediaPlayer1;
		private System.Windows.Forms.Button buttonOk;
		private System.Windows.Forms.Button buttonCancel;
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label labelStatus;
		private System.Windows.Forms.Button buttonBuildPrevious;
		private System.Windows.Forms.GroupBox groupBoxMarkOut;
		private System.Windows.Forms.GroupBox groupBoxMarkIn;
		private System.Windows.Forms.Button buttonIncAndBuild;

		#endregion Windows Forms Members

		#region Private Members

		private ArchiveTranscoderJobSegment segment;
		private DateTime markIn = DateTime.MinValue;
		private DateTime markOut = DateTime.MinValue;
		private ArchiveTranscoder transcoder = null;
		private bool buildingPreview = false;
		private DateTime currentPreviewStart = DateTime.MinValue;
		private String previewFilePath = null;
		private String previewDirPath = null;
		private String sqlHost;
		private String dbName;
		private bool closing;
		private StreamGroup videoStreamGroup;
        private bool useSlideStream;

		#endregion Private Members

		#region Properties

		public DateTime MarkIn
		{
			get {return markIn;}
		}
		public DateTime MarkOut
		{
			get {return markOut;}
		}

		#endregion Properties

		#region Constructor/Dispose

		public PreviewForm(ArchiveTranscoderJobSegment segment,String sqlHost,String dbName)
		{
			InitializeComponent();			
			this.sqlHost = sqlHost;
			this.dbName = dbName;
			this.segment = segment;
			this.textBoxStartTime.Text = segment.StartTime;
			this.textBoxDuration.Text = "300";
            useSlideStream = false;
            Debug.WriteLine("WMP Version: " + this.axWindowsMediaPlayer1.versionInfo); 
            if (segment.VideoDescriptor != null)
            {
                this.videoStreamGroup = new StreamGroup(segment.VideoDescriptor);
                this.labelVideoSource.Text = videoStreamGroup.ToString();
            }
            else
            {
                if (Utility.SegmentFlagIsSet(segment, SegmentFlags.SlidesReplaceVideo))
                {
                    this.videoStreamGroup = new StreamGroup(segment.PresentationDescriptor.PresentationCname, "Presentation Slides", "dynamicVideo");
                    useSlideStream = true;
                }
                else
                {
                    Debug.Assert(false);
                }
                this.labelVideoSource.Text = "[Presentation] " + videoStreamGroup.Cname;
            }

			bool itemChecked = false;
			closing = false;

			//We prefer to default to an audio cname that matches the video cname if possible.
			for (int i=0;i<segment.AudioDescriptor.Length;i++)
			{
				if (segment.AudioDescriptor[i].AudioCname == videoStreamGroup.Cname)
				{
					this.checkedListBoxAudioSources.Items.Add(new StreamGroup(segment.AudioDescriptor[i]),CheckState.Checked);
					itemChecked = true;
				}
				else
					this.checkedListBoxAudioSources.Items.Add(new StreamGroup(segment.AudioDescriptor[i]),CheckState.Unchecked);
			}
			//..otherwise just use the first one.
			if (!itemChecked)
			{
				if (this.checkedListBoxAudioSources.Items.Count >0)
				{
					this.checkedListBoxAudioSources.SetItemCheckState(0,CheckState.Checked);
				}
			}

			this.setToolTips();
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

		#endregion Constructor/Dispose

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PreviewForm));
            this.labelVideoSource = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.checkedListBoxAudioSources = new System.Windows.Forms.CheckedListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBoxStartDuration = new System.Windows.Forms.GroupBox();
            this.textBoxStartTime = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxDuration = new System.Windows.Forms.TextBox();
            this.buttonBuild = new System.Windows.Forms.Button();
            this.buttonMarkOut = new System.Windows.Forms.Button();
            this.buttonMarkIn = new System.Windows.Forms.Button();
            this.buttonOk = new System.Windows.Forms.Button();
            this.textBoxMarkIn = new System.Windows.Forms.TextBox();
            this.textBoxMarkOut = new System.Windows.Forms.TextBox();
            this.axWindowsMediaPlayer1 = new AxWMPLib.AxWindowsMediaPlayer();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.groupBoxMarkOut = new System.Windows.Forms.GroupBox();
            this.groupBoxMarkIn = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonIncAndBuild = new System.Windows.Forms.Button();
            this.buttonBuildPrevious = new System.Windows.Forms.Button();
            this.labelStatus = new System.Windows.Forms.Label();
            this.groupBoxStartDuration.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.axWindowsMediaPlayer1)).BeginInit();
            this.groupBoxMarkOut.SuspendLayout();
            this.groupBoxMarkIn.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelVideoSource
            // 
            this.labelVideoSource.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelVideoSource.Location = new System.Drawing.Point(96, 17);
            this.labelVideoSource.Name = "labelVideoSource";
            this.labelVideoSource.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.labelVideoSource.Size = new System.Drawing.Size(224, 13);
            this.labelVideoSource.TabIndex = 0;
            this.labelVideoSource.Text = "none";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(8, 24);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(64, 23);
            this.label2.TabIndex = 1;
            this.label2.Text = "Start Time:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // checkedListBoxAudioSources
            // 
            this.checkedListBoxAudioSources.CheckOnClick = true;
            this.checkedListBoxAudioSources.Location = new System.Drawing.Point(16, 56);
            this.checkedListBoxAudioSources.Name = "checkedListBoxAudioSources";
            this.checkedListBoxAudioSources.Size = new System.Drawing.Size(304, 64);
            this.checkedListBoxAudioSources.TabIndex = 2;
            this.checkedListBoxAudioSources.DoubleClick += new System.EventHandler(this.checkedListBoxAudioSources_DoubleClick);
            this.checkedListBoxAudioSources.SelectedIndexChanged += new System.EventHandler(this.checkedListBoxAudioSources_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(16, 40);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(200, 16);
            this.label3.TabIndex = 3;
            this.label3.Text = "Audio Source (Select One or More):";
            // 
            // groupBoxStartDuration
            // 
            this.groupBoxStartDuration.Controls.Add(this.textBoxStartTime);
            this.groupBoxStartDuration.Controls.Add(this.label2);
            this.groupBoxStartDuration.Controls.Add(this.label4);
            this.groupBoxStartDuration.Controls.Add(this.textBoxDuration);
            this.groupBoxStartDuration.Location = new System.Drawing.Point(16, 128);
            this.groupBoxStartDuration.Name = "groupBoxStartDuration";
            this.groupBoxStartDuration.Size = new System.Drawing.Size(304, 88);
            this.groupBoxStartDuration.TabIndex = 4;
            this.groupBoxStartDuration.TabStop = false;
            this.groupBoxStartDuration.Text = "Preview Start/Duration";
            // 
            // textBoxStartTime
            // 
            this.textBoxStartTime.Location = new System.Drawing.Point(72, 24);
            this.textBoxStartTime.Name = "textBoxStartTime";
            this.textBoxStartTime.Size = new System.Drawing.Size(224, 20);
            this.textBoxStartTime.TabIndex = 2;
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(8, 56);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(112, 23);
            this.label4.TabIndex = 1;
            this.label4.Text = "Duration (Seconds):";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // textBoxDuration
            // 
            this.textBoxDuration.Location = new System.Drawing.Point(152, 56);
            this.textBoxDuration.Name = "textBoxDuration";
            this.textBoxDuration.Size = new System.Drawing.Size(144, 20);
            this.textBoxDuration.TabIndex = 2;
            // 
            // buttonBuild
            // 
            this.buttonBuild.Location = new System.Drawing.Point(16, 224);
            this.buttonBuild.Name = "buttonBuild";
            this.buttonBuild.Size = new System.Drawing.Size(88, 24);
            this.buttonBuild.TabIndex = 5;
            this.buttonBuild.Text = "Build Preview";
            this.buttonBuild.Click += new System.EventHandler(this.buttonBuild_Click);
            // 
            // buttonMarkOut
            // 
            this.buttonMarkOut.Enabled = false;
            this.buttonMarkOut.Location = new System.Drawing.Point(8, 16);
            this.buttonMarkOut.Name = "buttonMarkOut";
            this.buttonMarkOut.Size = new System.Drawing.Size(48, 24);
            this.buttonMarkOut.TabIndex = 6;
            this.buttonMarkOut.Text = "Set";
            this.buttonMarkOut.Click += new System.EventHandler(this.buttonMarkOut_Click);
            // 
            // buttonMarkIn
            // 
            this.buttonMarkIn.Enabled = false;
            this.buttonMarkIn.Location = new System.Drawing.Point(8, 16);
            this.buttonMarkIn.Name = "buttonMarkIn";
            this.buttonMarkIn.Size = new System.Drawing.Size(48, 24);
            this.buttonMarkIn.TabIndex = 7;
            this.buttonMarkIn.Text = "Set";
            this.buttonMarkIn.Click += new System.EventHandler(this.buttonMarkIn_Click);
            // 
            // buttonOk
            // 
            this.buttonOk.Location = new System.Drawing.Point(608, 336);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(64, 24);
            this.buttonOk.TabIndex = 8;
            this.buttonOk.Text = "OK";
            this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
            // 
            // textBoxMarkIn
            // 
            this.textBoxMarkIn.Location = new System.Drawing.Point(64, 16);
            this.textBoxMarkIn.Name = "textBoxMarkIn";
            this.textBoxMarkIn.Size = new System.Drawing.Size(232, 20);
            this.textBoxMarkIn.TabIndex = 9;
            // 
            // textBoxMarkOut
            // 
            this.textBoxMarkOut.Location = new System.Drawing.Point(64, 16);
            this.textBoxMarkOut.Name = "textBoxMarkOut";
            this.textBoxMarkOut.Size = new System.Drawing.Size(232, 20);
            this.textBoxMarkOut.TabIndex = 9;
            // 
            // axWindowsMediaPlayer1
            // 
            this.axWindowsMediaPlayer1.Enabled = true;
            this.axWindowsMediaPlayer1.Location = new System.Drawing.Point(352, 8);
            this.axWindowsMediaPlayer1.Name = "axWindowsMediaPlayer1";
            this.axWindowsMediaPlayer1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axWindowsMediaPlayer1.OcxState")));
            this.axWindowsMediaPlayer1.Size = new System.Drawing.Size(320, 305);
            this.axWindowsMediaPlayer1.TabIndex = 10;
            // 
            // buttonCancel
            // 
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(520, 336);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(64, 24);
            this.buttonCancel.TabIndex = 11;
            this.buttonCancel.Text = "Cancel";
            // 
            // groupBoxMarkOut
            // 
            this.groupBoxMarkOut.Controls.Add(this.buttonMarkOut);
            this.groupBoxMarkOut.Controls.Add(this.textBoxMarkOut);
            this.groupBoxMarkOut.Location = new System.Drawing.Point(16, 336);
            this.groupBoxMarkOut.Name = "groupBoxMarkOut";
            this.groupBoxMarkOut.Size = new System.Drawing.Size(304, 48);
            this.groupBoxMarkOut.TabIndex = 12;
            this.groupBoxMarkOut.TabStop = false;
            this.groupBoxMarkOut.Text = "Segment Mark Out Time";
            // 
            // groupBoxMarkIn
            // 
            this.groupBoxMarkIn.Controls.Add(this.buttonMarkIn);
            this.groupBoxMarkIn.Controls.Add(this.textBoxMarkIn);
            this.groupBoxMarkIn.Location = new System.Drawing.Point(16, 280);
            this.groupBoxMarkIn.Name = "groupBoxMarkIn";
            this.groupBoxMarkIn.Size = new System.Drawing.Size(304, 48);
            this.groupBoxMarkIn.TabIndex = 0;
            this.groupBoxMarkIn.TabStop = false;
            this.groupBoxMarkIn.Text = "Segment Mark In Time";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(8, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 16);
            this.label1.TabIndex = 13;
            this.label1.Text = "Video Source:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // buttonIncAndBuild
            // 
            this.buttonIncAndBuild.Location = new System.Drawing.Point(240, 224);
            this.buttonIncAndBuild.Name = "buttonIncAndBuild";
            this.buttonIncAndBuild.Size = new System.Drawing.Size(80, 24);
            this.buttonIncAndBuild.TabIndex = 14;
            this.buttonIncAndBuild.Text = "Build Next >>";
            this.buttonIncAndBuild.Click += new System.EventHandler(this.buttonIncAndBuild_Click);
            // 
            // buttonBuildPrevious
            // 
            this.buttonBuildPrevious.Location = new System.Drawing.Point(152, 224);
            this.buttonBuildPrevious.Name = "buttonBuildPrevious";
            this.buttonBuildPrevious.Size = new System.Drawing.Size(80, 24);
            this.buttonBuildPrevious.TabIndex = 15;
            this.buttonBuildPrevious.Text = "<< Build Prev";
            this.buttonBuildPrevious.Click += new System.EventHandler(this.buttonBuildPrevious_Click);
            // 
            // labelStatus
            // 
            this.labelStatus.Location = new System.Drawing.Point(16, 256);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(304, 16);
            this.labelStatus.TabIndex = 16;
            // 
            // PreviewForm
            // 
            this.AcceptButton = this.buttonOk;
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.CancelButton = this.buttonCancel;
            this.ClientSize = new System.Drawing.Size(690, 392);
            this.Controls.Add(this.labelStatus);
            this.Controls.Add(this.buttonBuildPrevious);
            this.Controls.Add(this.buttonIncAndBuild);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBoxMarkOut);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.axWindowsMediaPlayer1);
            this.Controls.Add(this.buttonOk);
            this.Controls.Add(this.buttonBuild);
            this.Controls.Add(this.groupBoxStartDuration);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.checkedListBoxAudioSources);
            this.Controls.Add(this.labelVideoSource);
            this.Controls.Add(this.groupBoxMarkIn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PreviewForm";
            this.ShowInTaskbar = false;
            this.Text = "Windows Media Preview";
            this.Closing += new System.ComponentModel.CancelEventHandler(this.PreviewForm_Closing);
            this.groupBoxStartDuration.ResumeLayout(false);
            this.groupBoxStartDuration.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.axWindowsMediaPlayer1)).EndInit();
            this.groupBoxMarkOut.ResumeLayout(false);
            this.groupBoxMarkOut.PerformLayout();
            this.groupBoxMarkIn.ResumeLayout(false);
            this.groupBoxMarkIn.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		#region Form/Control Event Handlers

		private void buttonOk_Click(object sender, System.EventArgs e)
		{
			//Validate mark in/out times if any:
			if (this.textBoxMarkIn.Text.Trim() != "")
			{
				try
				{
					this.markIn = DateTime.Parse(this.textBoxMarkIn.Text.Trim());
				}
				catch
				{
					MessageBox.Show("Failed to parse Mark-In time as a date/time.","Error");
					this.markIn = DateTime.MinValue;
					return;
				}
			}
			else
			{
				this.markIn = DateTime.MinValue;
			}

			if (this.textBoxMarkOut.Text.Trim() != "")
			{
				try
				{
					this.markOut = DateTime.Parse(this.textBoxMarkOut.Text.Trim());
				}
				catch
				{
					MessageBox.Show("Failed to parse Mark-Out time as a date/time.","Error");
					this.markOut = DateTime.MinValue;
					return;
				}
			}
			else
			{
				this.markOut = DateTime.MinValue;
			}

			if ((markOut != DateTime.MinValue) && (markIn != DateTime.MinValue))
			{
				if (markOut < markIn)
				{
					MessageBox.Show("Mark-Out time must be greater than Mark-In time", "Error");
					return;
				}
			}

			this.DialogResult = DialogResult.OK;
			this.Close();
		}

		private void buttonMarkIn_Click(object sender, System.EventArgs e)
		{
			double currentPos = this.axWindowsMediaPlayer1.Ctlcontrols.currentPosition;
			DateTime markTime = currentPreviewStart + TimeSpan.FromSeconds(currentPos);
            //this.textBoxMarkIn.Text = markTime.ToString(Constants.dtformat);
            this.textBoxMarkIn.Text = Utility.GetLocalizedDateTimeString(markTime,Constants.timeformat);
        }

		private void buttonMarkOut_Click(object sender, System.EventArgs e)
		{
			double currentPos = this.axWindowsMediaPlayer1.Ctlcontrols.currentPosition;
			DateTime markTime = currentPreviewStart + TimeSpan.FromSeconds(currentPos);
            //this.textBoxMarkOut.Text = markTime.ToString(Constants.dtformat);
            this.textBoxMarkOut.Text = Utility.GetLocalizedDateTimeString(markTime, Constants.timeformat);
        }

		private void checkedListBoxAudioSources_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			//Always require at least one audio source to be checked.
			int selectedIndex = checkedListBoxAudioSources.SelectedIndex;
			if ((checkedListBoxAudioSources.CheckedItems.Count==0) && (!checkedListBoxAudioSources.GetItemChecked(selectedIndex)))
			{
				checkedListBoxAudioSources.SetItemChecked(selectedIndex,true);
			}
		}

		//Ignore double-clicks on the last checked item in the list box
		private void checkedListBoxAudioSources_DoubleClick(object sender, System.EventArgs e)
		{
			if (this.checkedListBoxAudioSources.SelectedIndex != -1)
				this.checkedListBoxAudioSources.SetItemChecked(this.checkedListBoxAudioSources.SelectedIndex,true);		
		}

		private void buttonIncAndBuild_Click(object sender, System.EventArgs e)
		{
			//Add duration to start time, then build the next chunk.
			//validate start time and duration.
			DateTime start = DateTime.MinValue;
			DateTime end = DateTime.MinValue;
			int duration = 0;
			try
			{
				start = DateTime.Parse(this.textBoxStartTime.Text);
			}
			catch
			{
				MessageBox.Show("Failed to parse Start Time as a date/time: " + this.textBoxStartTime.Text);
				return;
			}
			
			try
			{
				duration = Int32.Parse(this.textBoxDuration.Text);
			}
			catch
			{
				MessageBox.Show("Failed to parse preview duration as a positive integer.");
				return;
			}
            //this.textBoxStartTime.Text = ((DateTime)(start + TimeSpan.FromSeconds(duration))).ToString(Constants.dtformat);
            this.textBoxStartTime.Text = Utility.GetLocalizedDateTimeString((DateTime)(start + TimeSpan.FromSeconds(duration)), Constants.timeformat);
            this.buttonBuild_Click(sender, e);

		}

		private void buttonBuildPrevious_Click(object sender, System.EventArgs e)
		{
			//subtract duration from start time, then build.
			//validate start time and duration.
			DateTime start = DateTime.MinValue;
			DateTime end = DateTime.MinValue;
			int duration = 0;
			try
			{
				start = DateTime.Parse(this.textBoxStartTime.Text);
			}
			catch
			{
				MessageBox.Show("Failed to parse Start Time as a date/time: " + this.textBoxStartTime.Text);
				return;
			}
			
			try
			{
				duration = Int32.Parse(this.textBoxDuration.Text);
			}
			catch
			{
				MessageBox.Show("Failed to parse preview duration as a positive integer.");
				return;
			}
            //this.textBoxStartTime.Text = ((DateTime)(start - TimeSpan.FromSeconds(duration))).ToString(Constants.dtformat);
            this.textBoxStartTime.Text = Utility.GetLocalizedDateTimeString((DateTime)(start - TimeSpan.FromSeconds(duration)), Constants.timeformat);
            this.buttonBuild_Click(sender, e);		
		}

		private void buttonBuild_Click(object sender, System.EventArgs e)
		{
			if (! buildingPreview)
			{
				//validate start time and duration
				DateTime start = DateTime.MinValue;
				DateTime end = DateTime.MinValue;
				int duration = 0;
				try
				{
					start = DateTime.Parse(this.textBoxStartTime.Text);
				}
				catch
				{
					MessageBox.Show("Failed to parse Start Time as a date/time: " + this.textBoxStartTime.Text);
					return;
				}
			
				try
				{
					duration = Int32.Parse(this.textBoxDuration.Text);
				}
				catch
				{
					MessageBox.Show("Failed to parse preview duration as a positive integer.");
					return;
				}

				if (Directory.Exists(this.previewDirPath))
				{
					this.axWindowsMediaPlayer1.URL = "";
                    try {
                        Directory.Delete(this.previewDirPath, true);
                    }
                    catch (System.IO.IOException) { 
                        //This appears to be a bug.  Even with the recursive flag, it sometimes throws the 
                        //"Directory is not empty" exception.  Seems safe to ignore.
                        #if DEBUG
                        MessageBox.Show("Debug message: Directory delete operation may have failed for " + this.previewDirPath +
                            ".  It is most likely safe to ignore this message.");
                        #endif
                    }
				}

				end = start + TimeSpan.FromSeconds(duration);

				//create an encoder and load a job
				transcoder = new ArchiveTranscoder();
				transcoder.SQLHost = sqlHost;
				transcoder.DBName = dbName;
				ArchiveTranscoderBatch batch = new ArchiveTranscoderBatch();
				batch.Job = new ArchiveTranscoderJob[1];
				batch.Job[0] = new ArchiveTranscoderJob();
				batch.Job[0].ArchiveName = "ArchiveTranscoder_preview";
				batch.Job[0].Target = new ArchiveTranscoderJobTarget[1];
				batch.Job[0].Target[0] = new ArchiveTranscoderJobTarget();
				batch.Job[0].Target[0].Type = "stream";
				batch.Job[0].WMProfile = "norecompression";
				batch.Job[0].BaseName = "ArchiveTranscoder_preview";
				batch.Job[0].Path = Constants.TempPath;
				batch.Job[0].Segment = new ArchiveTranscoderJobSegment[1];
				batch.Job[0].Segment[0] = new ArchiveTranscoderJobSegment();
				batch.Job[0].Segment[0].VideoDescriptor = this.videoStreamGroup.ToVideoDescriptor();
				batch.Job[0].Segment[0].AudioDescriptor = new ArchiveTranscoderJobSegmentAudioDescriptor[this.checkedListBoxAudioSources.CheckedItems.Count];
				int i = 0;
				foreach (StreamGroup sg in this.checkedListBoxAudioSources.CheckedItems)
				{
					batch.Job[0].Segment[0].AudioDescriptor[i] = sg.ToAudioDescriptor();
					i++;
				}
                //batch.Job[0].Segment[0].StartTime = start.ToString(Constants.dtformat);
                batch.Job[0].Segment[0].StartTime = Utility.GetLocalizedDateTimeString(start, Constants.timeformat);
                //batch.Job[0].Segment[0].EndTime = end.ToString(Constants.dtformat);
                batch.Job[0].Segment[0].EndTime = Utility.GetLocalizedDateTimeString(end, Constants.timeformat);

                if (useSlideStream)
                {
                    Utility.SetSegmentFlag(batch.Job[0].Segment[0], SegmentFlags.SlidesReplaceVideo);
                    //Since slide streams start out uncompressed, we can't use norecompression.
                    //Use the 256kbps system profile
				    batch.Job[0].WMProfile = "9";
                    //In this case we also need a presentation source and slide decks
                    batch.Job[0].Segment[0].PresentationDescriptor = segment.PresentationDescriptor;
                    batch.Job[0].Segment[0].SlideDecks = segment.SlideDecks;
                }

				transcoder.LoadJobs(batch);

				previewDirPath = Path.Combine(Constants.TempPath,@"ArchiveTranscoder_preview");
				previewFilePath = Path.Combine(previewDirPath,@"stream\ArchiveTranscoder_preview.wmv");

				//register for stop encode and status events
				transcoder.OnBatchCompleted += new ArchiveTranscoder.batchCompletedHandler(OnBuildCompleted);
				transcoder.OnStatusReport += new ArchiveTranscoder.statusReportHandler(OnStatusReport);
				transcoder.OnStopCompleted += new ArchiveTranscoder.stopCompletedHandler(StopCompletedHandler);

				//disable controls during build
				BuildControlsEnable(false,false);

				String result = transcoder.Encode();

				if (result == null)
				{
					buttonBuild.Text = "Stop Build";
					buildingPreview = true;
					this.currentPreviewStart = start;
				}
				else
				{
					MessageBox.Show("Preview build failed to start: " + result);
					BuildControlsEnable(true,false);
				}
			}
			else
			{
				this.Cursor = Cursors.WaitCursor;
				//stop build in progress..
				if (transcoder != null)
					transcoder.AsyncStop();

				this.buttonBuild.Enabled = false;
				this.buttonBuild.Text = "Stopping ...";
			}
		}


		/// <summary>
		/// If the user clicks the X in the upper right of the form, this is where we stop and clean up
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void PreviewForm_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			this.axWindowsMediaPlayer1.URL = "";

			closing = true;

			if (transcoder != null)
			{
				transcoder.Stop();
			}

			if (Directory.Exists(previewDirPath))
			{
				try
				{
					Directory.Delete(previewDirPath,true);
				}
				catch (Exception ex)
				{
					Debug.WriteLine("failed to delete preview directory: " + ex.ToString());
				}
			}
		}

		#endregion Form/Control Event Handlers

		#region Transcoder Event Handlers

		private void OnBuildCompleted()
		{
			if (closing)
				return;

			if (this.InvokeRequired)
			{
				this.Invoke(new BuildCompletedDelegate(BuildCompleted));
			}
			else
			{
				BuildCompleted();
			}
		}

		private delegate void BuildCompletedDelegate();
		private void BuildCompleted()
		{
			buildingPreview = false;
			this.buttonBuild.Text = "Build Preview";

			//re-enable controls
			BuildControlsEnable(true,true);
			this.buttonMarkOut.Enabled = true;
			this.buttonMarkIn.Enabled = true;

			if (transcoder.BatchErrorLevel >= 5)
			{
				MessageBox.Show(this,"There were problems building the preview: \r\n" + transcoder.BatchLogString,
					"Problems Building Preview",MessageBoxButtons.OK,MessageBoxIcon.Warning);
			}
			if (transcoder.BatchErrorLevel < 6)
			{
				//load result into WMP;
				if (File.Exists(previewFilePath))
				{
					this.axWindowsMediaPlayer1.URL = previewFilePath;
				}
			}
			transcoder = null;
			this.Cursor = Cursors.Default;

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
					Debug.WriteLine("PreviewForm.StopCompletedHandler couldn't invoke: " + e.ToString());
				}
			}
			else
			{
				StopCompleted();
			}
		}

		private void StopCompleted()
		{
			transcoder = null;
			buttonBuild.Text = "Build Preview";
			buttonBuild.Enabled = true;
			buildingPreview = false;
			BuildControlsEnable(true,false);
			this.Cursor = Cursors.Default;
		}

		#endregion Transcoder Event Handlers

		#region Private Methods

		/// <summary>
		/// Call with enable==false at the start of a build and enable==true at the end.
		/// At the end of a build the success flag is also observed: if the build completed
		/// successfully, success is true.
		/// </summary>
		/// <param name="enable"></param>
		/// <param name="success"></param>
		private void BuildControlsEnable(bool enable, bool success)
		{
			this.checkedListBoxAudioSources.Enabled = enable;
			this.textBoxStartTime.Enabled = enable;
			this.textBoxDuration.Enabled = enable;
			this.buttonOk.Enabled = enable;
			this.buttonCancel.Enabled = enable;
			this.buttonIncAndBuild.Enabled = enable;
			this.buttonBuildPrevious.Enabled = enable;
			if (enable)
			{
				this.buttonMarkOut.Enabled = success;
				this.buttonMarkIn.Enabled = success;
			}
			else
			{
				this.buttonMarkOut.Enabled = false;
				this.buttonMarkIn.Enabled = false;
			}
		}

		private void setToolTips()
		{
			ToolTip tt = new ToolTip();
			tt.SetToolTip(this.checkedListBoxAudioSources,"Note: It is faster to build a preview with only one audio source selected.");
			tt.SetToolTip(this.buttonBuildPrevious,"Subtract duration from start time, then build the preview.");
			tt.SetToolTip(this.buttonIncAndBuild,"Add duration to start time, then build the preview.");
			tt.SetToolTip(this.textBoxMarkOut,"If set, this time will replace the segment end time.");
			tt.SetToolTip(this.textBoxMarkIn,"If set, this time will replace the segment start time.");
		}

		#endregion Private Methods

	}
}
