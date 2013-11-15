using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Diagnostics;
using System.Text;
using MSR.LST.Net.Rtp;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Allow the user to build an archive segment from the database.
	/// 
	/// By definition a segment consists of one time range, one video cname and one or more audio cnames.  For each cname
	/// there may be one or more non-overlapping streams.  There also may be an optional Presentation Cname.
	/// There is a mismatch between the abstraction which we want to present to the user, and the reality 
	/// of CXP reflected in the database schema.  In particular, the user should not need to care whether a node rejoined the
	/// venue during the timespan of interest, so we hide streams, and present only cname and payload.
	/// 
	/// We make an assumption that all streams in a segment are in the same conference (Ie. they have the same SQL Conference ID).  
	/// This means that if the archiver needed to be stopped and restarted during a conference, we would need to use 
	/// multiple segments to build the Windows Media archive.
	/// This also means that all sources including presentation were in the same venue when recorded.
	/// </summary>
	public class SegmentForm : System.Windows.Forms.Form
	{
		#region Windows Forms Members
		private System.Windows.Forms.TreeView treeViewSources;
		private System.Windows.Forms.ListBox listBoxAudioSources;
		private System.Windows.Forms.Button buttonRemoveAudio;
		private System.Windows.Forms.Button buttonPreview;
		private System.Windows.Forms.Button buttonCancel;
		private System.Windows.Forms.Button buttonOk;
		private System.Windows.Forms.GroupBox groupBoxAudio;
		private System.Windows.Forms.GroupBox groupBoxVideo;
		private System.Windows.Forms.GroupBox groupBoxPresentation;
		private System.Windows.Forms.GroupBox groupBoxStartEnd;
		private System.Windows.Forms.TextBox textBoxEnd;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox textBoxStart;
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.ListBox listBoxPresentationSource;
		private System.Windows.Forms.ListBox listBoxVideoSource;
		private System.Windows.Forms.GroupBox groupBoxConference;
		private System.Windows.Forms.ListBox listBoxConference;
		private System.Windows.Forms.Button buttonDeckMatcher;
		private System.Windows.Forms.Button buttonRemovePresentation;
        private CheckBox checkBoxSlidesReplaceVideo;

		#endregion Windows Forms Members
		
		#region Private Members

		private Deck[] slideDecks;
		private String sqlHost;
		private String dbName;
		private ArchiveTranscoderJobSegment segment;
        private Button buttonPresentationFromVideo;
		private ConferenceDataCache confDataCache;

		#endregion Private Members

		#region Public Properties

		public ArchiveTranscoderJobSegment Segment
		{
			get {return segment;}
		}

		#endregion Public Properties

		#region Construct/Dispose

		public SegmentForm(ConferenceDataCache confDataCache)
		{
			InitializeComponent();
			this.setToolTips();
			segment = null;
			slideDecks = null;
			this.confDataCache = confDataCache;
			this.sqlHost = confDataCache.SqlServer;
			this.dbName = confDataCache.DbName;
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

		#endregion Construct/Dispose

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.treeViewSources = new System.Windows.Forms.TreeView();
            this.listBoxAudioSources = new System.Windows.Forms.ListBox();
            this.buttonRemoveAudio = new System.Windows.Forms.Button();
            this.buttonPreview = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonOk = new System.Windows.Forms.Button();
            this.groupBoxAudio = new System.Windows.Forms.GroupBox();
            this.groupBoxVideo = new System.Windows.Forms.GroupBox();
            this.listBoxVideoSource = new System.Windows.Forms.ListBox();
            this.groupBoxPresentation = new System.Windows.Forms.GroupBox();
            this.checkBoxSlidesReplaceVideo = new System.Windows.Forms.CheckBox();
            this.listBoxPresentationSource = new System.Windows.Forms.ListBox();
            this.buttonDeckMatcher = new System.Windows.Forms.Button();
            this.buttonRemovePresentation = new System.Windows.Forms.Button();
            this.groupBoxStartEnd = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxEnd = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxStart = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBoxConference = new System.Windows.Forms.GroupBox();
            this.listBoxConference = new System.Windows.Forms.ListBox();
            this.buttonPresentationFromVideo = new System.Windows.Forms.Button();
            this.groupBoxAudio.SuspendLayout();
            this.groupBoxVideo.SuspendLayout();
            this.groupBoxPresentation.SuspendLayout();
            this.groupBoxStartEnd.SuspendLayout();
            this.groupBoxConference.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeViewSources
            // 
            this.treeViewSources.Location = new System.Drawing.Point(16, 40);
            this.treeViewSources.Name = "treeViewSources";
            this.treeViewSources.Size = new System.Drawing.Size(344, 557);
            this.treeViewSources.TabIndex = 0;
            this.treeViewSources.DoubleClick += new System.EventHandler(this.treeViewSources_DoubleClick);
            // 
            // listBoxAudioSources
            // 
            this.listBoxAudioSources.Location = new System.Drawing.Point(376, 144);
            this.listBoxAudioSources.Name = "listBoxAudioSources";
            this.listBoxAudioSources.Size = new System.Drawing.Size(304, 95);
            this.listBoxAudioSources.TabIndex = 1;
            this.listBoxAudioSources.SelectedIndexChanged += new System.EventHandler(this.listBoxAudioSources_SelectedIndexChanged);
            // 
            // buttonRemoveAudio
            // 
            this.buttonRemoveAudio.Enabled = false;
            this.buttonRemoveAudio.Location = new System.Drawing.Point(248, 16);
            this.buttonRemoveAudio.Name = "buttonRemoveAudio";
            this.buttonRemoveAudio.Size = new System.Drawing.Size(64, 24);
            this.buttonRemoveAudio.TabIndex = 2;
            this.buttonRemoveAudio.Text = "Remove";
            this.buttonRemoveAudio.Click += new System.EventHandler(this.buttonRemoveAudio_Click);
            // 
            // buttonPreview
            // 
            this.buttonPreview.Enabled = false;
            this.buttonPreview.Location = new System.Drawing.Point(432, 315);
            this.buttonPreview.Name = "buttonPreview";
            this.buttonPreview.Size = new System.Drawing.Size(176, 24);
            this.buttonPreview.TabIndex = 2;
            this.buttonPreview.Text = "Audio/Video Preview";
            this.buttonPreview.Click += new System.EventHandler(this.buttonPreview_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(504, 574);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(72, 24);
            this.buttonCancel.TabIndex = 2;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // buttonOk
            // 
            this.buttonOk.Enabled = false;
            this.buttonOk.Location = new System.Drawing.Point(608, 574);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(72, 24);
            this.buttonOk.TabIndex = 2;
            this.buttonOk.Text = "OK";
            this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
            // 
            // groupBoxAudio
            // 
            this.groupBoxAudio.Controls.Add(this.buttonRemoveAudio);
            this.groupBoxAudio.Location = new System.Drawing.Point(368, 96);
            this.groupBoxAudio.Name = "groupBoxAudio";
            this.groupBoxAudio.Size = new System.Drawing.Size(320, 152);
            this.groupBoxAudio.TabIndex = 3;
            this.groupBoxAudio.TabStop = false;
            this.groupBoxAudio.Text = "Selected Audio Sources";
            // 
            // groupBoxVideo
            // 
            this.groupBoxVideo.Controls.Add(this.listBoxVideoSource);
            this.groupBoxVideo.Location = new System.Drawing.Point(368, 256);
            this.groupBoxVideo.Name = "groupBoxVideo";
            this.groupBoxVideo.Size = new System.Drawing.Size(320, 48);
            this.groupBoxVideo.TabIndex = 4;
            this.groupBoxVideo.TabStop = false;
            this.groupBoxVideo.Text = "Selected Video Source";
            // 
            // listBoxVideoSource
            // 
            this.listBoxVideoSource.Location = new System.Drawing.Point(8, 16);
            this.listBoxVideoSource.Name = "listBoxVideoSource";
            this.listBoxVideoSource.Size = new System.Drawing.Size(304, 17);
            this.listBoxVideoSource.TabIndex = 1;
            // 
            // groupBoxPresentation
            // 
            this.groupBoxPresentation.Controls.Add(this.buttonPresentationFromVideo);
            this.groupBoxPresentation.Controls.Add(this.checkBoxSlidesReplaceVideo);
            this.groupBoxPresentation.Controls.Add(this.listBoxPresentationSource);
            this.groupBoxPresentation.Controls.Add(this.buttonDeckMatcher);
            this.groupBoxPresentation.Controls.Add(this.buttonRemovePresentation);
            this.groupBoxPresentation.Location = new System.Drawing.Point(368, 350);
            this.groupBoxPresentation.Name = "groupBoxPresentation";
            this.groupBoxPresentation.Size = new System.Drawing.Size(320, 126);
            this.groupBoxPresentation.TabIndex = 5;
            this.groupBoxPresentation.TabStop = false;
            this.groupBoxPresentation.Text = "Selected Presentation Source (Optional)";
            // 
            // checkBoxSlidesReplaceVideo
            // 
            this.checkBoxSlidesReplaceVideo.AutoSize = true;
            this.checkBoxSlidesReplaceVideo.Enabled = false;
            this.checkBoxSlidesReplaceVideo.Location = new System.Drawing.Point(26, 73);
            this.checkBoxSlidesReplaceVideo.Name = "checkBoxSlidesReplaceVideo";
            this.checkBoxSlidesReplaceVideo.Size = new System.Drawing.Size(175, 17);
            this.checkBoxSlidesReplaceVideo.TabIndex = 9;
            this.checkBoxSlidesReplaceVideo.Text = "Create Video From Presentation";
            this.checkBoxSlidesReplaceVideo.UseVisualStyleBackColor = true;
            this.checkBoxSlidesReplaceVideo.CheckedChanged += new System.EventHandler(this.checkBoxSlidesReplaceVideo_CheckedChanged);
            // 
            // listBoxPresentationSource
            // 
            this.listBoxPresentationSource.Location = new System.Drawing.Point(8, 16);
            this.listBoxPresentationSource.Name = "listBoxPresentationSource";
            this.listBoxPresentationSource.Size = new System.Drawing.Size(304, 17);
            this.listBoxPresentationSource.TabIndex = 0;
            // 
            // buttonDeckMatcher
            // 
            this.buttonDeckMatcher.Enabled = false;
            this.buttonDeckMatcher.Location = new System.Drawing.Point(24, 42);
            this.buttonDeckMatcher.Name = "buttonDeckMatcher";
            this.buttonDeckMatcher.Size = new System.Drawing.Size(184, 23);
            this.buttonDeckMatcher.TabIndex = 8;
            this.buttonDeckMatcher.Text = "Specify Presentation Decks";
            this.buttonDeckMatcher.Click += new System.EventHandler(this.buttonDeckMatcher_Click);
            // 
            // buttonRemovePresentation
            // 
            this.buttonRemovePresentation.Enabled = false;
            this.buttonRemovePresentation.Location = new System.Drawing.Point(224, 42);
            this.buttonRemovePresentation.Name = "buttonRemovePresentation";
            this.buttonRemovePresentation.Size = new System.Drawing.Size(64, 24);
            this.buttonRemovePresentation.TabIndex = 3;
            this.buttonRemovePresentation.Text = "Remove";
            this.buttonRemovePresentation.Click += new System.EventHandler(this.buttonRemovePresentation_Click);
            // 
            // groupBoxStartEnd
            // 
            this.groupBoxStartEnd.Controls.Add(this.label1);
            this.groupBoxStartEnd.Controls.Add(this.textBoxEnd);
            this.groupBoxStartEnd.Controls.Add(this.label2);
            this.groupBoxStartEnd.Controls.Add(this.textBoxStart);
            this.groupBoxStartEnd.Location = new System.Drawing.Point(368, 482);
            this.groupBoxStartEnd.Name = "groupBoxStartEnd";
            this.groupBoxStartEnd.Size = new System.Drawing.Size(248, 80);
            this.groupBoxStartEnd.TabIndex = 6;
            this.groupBoxStartEnd.TabStop = false;
            this.groupBoxStartEnd.Text = "Segment Start/End Time";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(16, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(32, 16);
            this.label1.TabIndex = 1;
            this.label1.Text = "Start:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // textBoxEnd
            // 
            this.textBoxEnd.Location = new System.Drawing.Point(56, 48);
            this.textBoxEnd.Name = "textBoxEnd";
            this.textBoxEnd.Size = new System.Drawing.Size(176, 20);
            this.textBoxEnd.TabIndex = 0;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(16, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(32, 16);
            this.label2.TabIndex = 1;
            this.label2.Text = "End:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // textBoxStart
            // 
            this.textBoxStart.Location = new System.Drawing.Point(56, 24);
            this.textBoxStart.Name = "textBoxStart";
            this.textBoxStart.Size = new System.Drawing.Size(176, 20);
            this.textBoxStart.TabIndex = 0;
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(16, 8);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(632, 23);
            this.label3.TabIndex = 7;
            this.label3.Text = "Double click on Audio, Video and Presentation streams in the tree below to add th" +
                "em to the archive segment.";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // groupBoxConference
            // 
            this.groupBoxConference.Controls.Add(this.listBoxConference);
            this.groupBoxConference.Location = new System.Drawing.Point(368, 40);
            this.groupBoxConference.Name = "groupBoxConference";
            this.groupBoxConference.Size = new System.Drawing.Size(320, 48);
            this.groupBoxConference.TabIndex = 5;
            this.groupBoxConference.TabStop = false;
            this.groupBoxConference.Text = "Selected Conference";
            // 
            // listBoxConference
            // 
            this.listBoxConference.Location = new System.Drawing.Point(8, 16);
            this.listBoxConference.Name = "listBoxConference";
            this.listBoxConference.Size = new System.Drawing.Size(304, 17);
            this.listBoxConference.TabIndex = 1;
            // 
            // buttonPresentationFromVideo
            // 
            this.buttonPresentationFromVideo.Enabled = true;
            this.buttonPresentationFromVideo.Location = new System.Drawing.Point(26, 96);
            this.buttonPresentationFromVideo.Name = "buttonPresentationFromVideo";
            this.buttonPresentationFromVideo.Size = new System.Drawing.Size(182, 23);
            this.buttonPresentationFromVideo.TabIndex = 10;
            this.buttonPresentationFromVideo.Text = "Create Presentation from Video ...";
            this.buttonPresentationFromVideo.Click += new System.EventHandler(this.buttonPresentationFromVideo_Click);
            // 
            // SegmentForm
            // 
            this.AcceptButton = this.buttonOk;
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.CancelButton = this.buttonCancel;
            this.ClientSize = new System.Drawing.Size(696, 609);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.groupBoxStartEnd);
            this.Controls.Add(this.listBoxAudioSources);
            this.Controls.Add(this.treeViewSources);
            this.Controls.Add(this.buttonPreview);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOk);
            this.Controls.Add(this.groupBoxAudio);
            this.Controls.Add(this.groupBoxVideo);
            this.Controls.Add(this.groupBoxPresentation);
            this.Controls.Add(this.groupBoxConference);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SegmentForm";
            this.ShowInTaskbar = false;
            this.Text = "Create/Edit Job Segment";
            this.groupBoxAudio.ResumeLayout(false);
            this.groupBoxVideo.ResumeLayout(false);
            this.groupBoxPresentation.ResumeLayout(false);
            this.groupBoxPresentation.PerformLayout();
            this.groupBoxStartEnd.ResumeLayout(false);
            this.groupBoxStartEnd.PerformLayout();
            this.groupBoxConference.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		#region Public Methods

		/// <summary>
		/// Load an existing segment for editing.
		/// Return true if the segment appears to match a conference in the database.
		/// </summary>
		/// <param name="segment"></param>
		public bool LoadSegment(ArchiveTranscoderJobSegment segment)
		{
			//fill in the form controls with existing data
			this.segment = segment;
			if ((segment.VideoDescriptor != null) && 
				(segment.VideoDescriptor.VideoCname != null) &&
				(segment.VideoDescriptor.VideoCname != ""))
			{
				this.listBoxVideoSource.Items.Add(new StreamGroup(segment.VideoDescriptor)); 
			}
            if ((segment.PresentationDescriptor != null) &&
                (segment.PresentationDescriptor.PresentationCname != null) &&
                (segment.PresentationDescriptor.PresentationCname != ""))
            {
                if (segment.PresentationDescriptor.VideoDescriptor == null) {
                    this.listBoxPresentationSource.Items.Add(new StreamGroup(segment.PresentationDescriptor));
                    this.buttonRemovePresentation.Enabled = true;
                    this.buttonDeckMatcher.Enabled = true;
                    if (segment.SlideDecks != null) {
                        slideDecks = new Deck[segment.SlideDecks.Length];
                        for (int i = 0; i < segment.SlideDecks.Length; i++) {
                            slideDecks[i] = new Deck(segment.SlideDecks[i]);
                        }
                    }
                    this.checkBoxSlidesReplaceVideo.Enabled = true;
                    if (Utility.SegmentFlagIsSet(segment, SegmentFlags.SlidesReplaceVideo)) {
                        this.checkBoxSlidesReplaceVideo.Checked = true;
                        //Other config is done by the check changed handler.
                    }
                }
                else {
                    //This is for building presentation from video stills.
                    this.listBoxPresentationSource.Items.Add(new StreamGroup(segment.PresentationDescriptor.VideoDescriptor));
                    this.buttonRemovePresentation.Enabled = false;
                    this.buttonDeckMatcher.Enabled = false;
                    this.checkBoxSlidesReplaceVideo.Enabled = false;
                    this.buttonRemovePresentation.Enabled = true;
                }
            }
            else
            {
                this.checkBoxSlidesReplaceVideo.Enabled = false;
            }
			if ((segment.AudioDescriptor != null) && (segment.AudioDescriptor.Length > 0))
			{
				for (int i=0; i< segment.AudioDescriptor.Length; i++)
				{
					if ((segment.AudioDescriptor[i].AudioCname != null) && (segment.AudioDescriptor[i].AudioCname != ""))
					{
						this.listBoxAudioSources.Items.Add(new StreamGroup(segment.AudioDescriptor[i]));
					}
				}
			}
			this.textBoxStart.Text = segment.StartTime;
			this.textBoxEnd.Text = segment.EndTime;

			String confStr = GetCurrentConfString(segment);
			if (confStr != "")
			{
				this.listBoxConference.Items.Add(confStr);
			}
			else
			{
				this.listBoxConference.Items.Add("Unavailable Conference!");
			}
			
			if ((this.listBoxAudioSources.Items.Count > 0) && (this.listBoxVideoSource.Items.Count>0))
			{
				this.buttonPreview.Enabled = true;
				this.buttonOk.Enabled = true;
			}

			if (confStr=="")
				return false;
			return true;
		}

		/// <summary>
		/// Walk the ConferenceDataCache result tree, building the TreeView.  Note that we assume the lookup is 
		/// done before we are invoked.
		/// </summary>
		public void Initialize()
		{
			foreach (ConferenceWrapper c in confDataCache.ConferenceData)
			{
				TreeNode newNode = new TreeNode("Conference: " + c.Conference.Start.ToString() + " - " + c.Conference.Description);
				newNode.Tag = c.Conference;
				foreach( ParticipantWrapper part in c.Participants )
				{
					// Create Participant node

					TreeNode partNode = new TreeNode("Participant: " + makeFriendlyCname(part.Participant.Name,part.Participant.CName));
					partNode.Tag = part.Participant;

					foreach (StreamGroup sg in part.StreamGrouper.GetStreamList())
					{
						TreeNode streamNode = partNode.Nodes.Add(sg.ToTreeViewString());
						streamNode.Tag = sg;
					}

					newNode.Nodes.Add(partNode);
				}
				this.treeViewSources.Nodes.Add(newNode);
			}
		}

		#endregion Public Methods

		#region UI Control Event Handlers

		private void buttonCancel_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.Cancel;
		}

		private void buttonOk_Click(object sender, System.EventArgs e)
		{
			DateTime start = DateTime.MinValue;
			DateTime end = DateTime.MinValue;

			if (!validate(out start, out end))
			{
				return;
			}

			//Warn if slide decks don't exist or are not matched.
            if (!validateDecks(sender, e))
            {
                return;
            }

			segment = new ArchiveTranscoderJobSegment();
            //segment.StartTime = start.ToString(Constants.dtformat);
            segment.StartTime = Utility.GetLocalizedDateTimeString(start, Constants.timeformat);
            //segment.EndTime = end.ToString(Constants.dtformat);
            segment.EndTime = Utility.GetLocalizedDateTimeString(end, Constants.timeformat);
            if (this.listBoxVideoSource.Items.Count == 1)
			{
                if (this.listBoxVideoSource.Items[0] is StreamGroup)
                    segment.VideoDescriptor = ((StreamGroup)this.listBoxVideoSource.Items[0]).ToVideoDescriptor();
                else
                    segment.VideoDescriptor = null; 
			}
			segment.AudioDescriptor = new ArchiveTranscoderJobSegmentAudioDescriptor[this.listBoxAudioSources.Items.Count];
			for (int i = 0; i < this.listBoxAudioSources.Items.Count; i++)
			{
				segment.AudioDescriptor[i] = ((StreamGroup)this.listBoxAudioSources.Items[i]).ToAudioDescriptor();
			}
			if (this.listBoxPresentationSource.Items.Count == 1)
			{
				segment.PresentationDescriptor = ((StreamGroup)this.listBoxPresentationSource.Items[0]).ToPresentationDescriptor();
				if (slideDecks != null)
				{
					segment.SlideDecks = new ArchiveTranscoderJobSlideDeck[slideDecks.Length];
					for (int i=0; i<slideDecks.Length; i++)
					{
						segment.SlideDecks[i] = new ArchiveTranscoderJobSlideDeck();
						segment.SlideDecks[i].DeckGuid = slideDecks[i].DeckGuid.ToString();
						segment.SlideDecks[i].Title = slideDecks[i].FileName;
						if (slideDecks[i].Matched)
							segment.SlideDecks[i].Path = slideDecks[i].Path;
					}
				}
                if (checkBoxSlidesReplaceVideo.Checked)
                {
                    Utility.SetSegmentFlag(segment, SegmentFlags.SlidesReplaceVideo);
                }

			}

			this.DialogResult = DialogResult.OK;
			this.Close();
		}

		private void treeViewSources_DoubleClick(object sender, System.EventArgs e)
		{
			if (sender is TreeView)
			{
				TreeView tv = (TreeView)sender;
				TreeNode tn = tv.SelectedNode;
				if ((tn != null) && (tn.Tag is StreamGroup)) //<--lame
				{
					Participant p = (Participant)tn.Parent.Tag;
					Conference c = (Conference)tn.Parent.Parent.Tag;
					StreamGroup s = (StreamGroup)tn.Tag;

					//changed conference.. reset some things:
					if (!listBoxConference.Items.Contains(c.Start.ToString() + " - " + c.Description))
					{
						listBoxConference.Items.Clear();
						listBoxConference.Items.Add(c.Start.ToString() + " - " + c.Description);
						listBoxVideoSource.Items.Clear();
						listBoxAudioSources.Items.Clear();
						listBoxPresentationSource.Items.Clear();

						//default time range is the entire conference
                        //textBoxStart.Text = c.Start.ToString(Constants.dtformat);
                        textBoxStart.Text = Utility.GetLocalizedDateTimeString(c.Start, Constants.timeformat);
                        //textBoxEnd.Text = c.End.ToString(Constants.dtformat);
                        textBoxEnd.Text = Utility.GetLocalizedDateTimeString(c.End, Constants.timeformat);
                        this.buttonRemoveAudio.Enabled = false;
					}

					if (s.Payload == "dynamicVideo")
					{
                        if (!this.checkBoxSlidesReplaceVideo.Checked)
                        {
                            this.listBoxVideoSource.Items.Clear();
                            this.listBoxVideoSource.Items.Add(s);
                        }
					}
					else if (s.Payload == "dynamicAudio")
					{
						if (!listBoxAudioSourcesContains(s))
							this.listBoxAudioSources.Items.Add(s);
					}
					else if ((s.Payload == "dynamicPresentation") || (s.Payload == "RTDocument"))
					{
						this.listBoxPresentationSource.Items.Clear();
						this.listBoxPresentationSource.Items.Add(s);
						this.buttonDeckMatcher.Enabled = true;
						this.buttonRemovePresentation.Enabled = true;
                        this.checkBoxSlidesReplaceVideo.Enabled = true;
					}

					if ((this.listBoxAudioSources.Items.Count > 0) && (this.listBoxVideoSource.Items.Count>0))
					{
						this.buttonPreview.Enabled = true;
						this.buttonOk.Enabled = true;
					}

				}
			}
		}

        /// <summary>
        /// Wrote a custom 'contains' method here because it may contain a different object with
        /// the same identifiers if it was preloaded.
        /// </summary>
        /// <param name="sg_target"></param>
        /// <returns></returns>
        private bool listBoxAudioSourcesContains(StreamGroup sg_target)
        {
            foreach (StreamGroup sg in listBoxAudioSources.Items)
            {
                if ((sg.Cname == sg_target.Cname) &&
                    (sg.Name == sg_target.Name))
                    return true;
            }
            return false;
        }


		private void listBoxAudioSources_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.buttonRemoveAudio.Enabled = true;
		}

		private void buttonRemoveAudio_Click(object sender, System.EventArgs e)
		{
			if (this.listBoxAudioSources.SelectedItem != null)
			{
				this.listBoxAudioSources.Items.Remove(this.listBoxAudioSources.SelectedItem);
				buttonRemoveAudio.Enabled = false;
				if (this.listBoxAudioSources.Items.Count == 0)
				{
					this.buttonPreview.Enabled = false;
					this.buttonOk.Enabled = false;
				}
			}
		}


		private void buttonPreview_Click(object sender, System.EventArgs e)
		{
			//This button should not be enabled unless there is at least one audio, one video, and a time range.
			
			DateTime start = DateTime.MinValue;
			DateTime end = DateTime.MinValue;

			if (!validate(out start, out end))
			{
				return;
			}

			segment = new ArchiveTranscoderJobSegment();
            //segment.StartTime = start.ToString(Constants.dtformat);
            //segment.EndTime = end.ToString(Constants.dtformat);
            segment.StartTime = Utility.GetLocalizedDateTimeString(start, Constants.timeformat);
            segment.EndTime = Utility.GetLocalizedDateTimeString(end, Constants.timeformat);
            if (this.listBoxVideoSource.Items[0] is StreamGroup)
			    segment.VideoDescriptor = ((StreamGroup)this.listBoxVideoSource.Items[0]).ToVideoDescriptor();

			segment.AudioDescriptor = new ArchiveTranscoderJobSegmentAudioDescriptor[this.listBoxAudioSources.Items.Count];
			for (int i = 0; i < this.listBoxAudioSources.Items.Count; i++)
			{
				segment.AudioDescriptor[i] = ((StreamGroup)this.listBoxAudioSources.Items[i]).ToAudioDescriptor();
			}

            if ((this.listBoxPresentationSource.Items.Count == 1) && (checkBoxSlidesReplaceVideo.Checked))
            {
                if (!validateDecks(sender,e))
                    return;

                segment.PresentationDescriptor = ((StreamGroup)this.listBoxPresentationSource.Items[0]).ToPresentationDescriptor();
                if (slideDecks != null)
                {
                    segment.SlideDecks = new ArchiveTranscoderJobSlideDeck[slideDecks.Length];
                    for (int i = 0; i < slideDecks.Length; i++)
                    {
                        segment.SlideDecks[i] = new ArchiveTranscoderJobSlideDeck();
                        segment.SlideDecks[i].DeckGuid = slideDecks[i].DeckGuid.ToString();
                        segment.SlideDecks[i].Title = slideDecks[i].FileName;
                        if (slideDecks[i].Matched)
                            segment.SlideDecks[i].Path = slideDecks[i].Path;
                    }
                }
                Utility.SetSegmentFlag(segment, SegmentFlags.SlidesReplaceVideo);
            }


			this.Cursor = Cursors.WaitCursor;
			PreviewForm previewForm = new PreviewForm(segment,sqlHost,dbName);
			this.Cursor = Cursors.Default;

			if (DialogResult.OK == previewForm.ShowDialog(this))
			{
				if (previewForm.MarkIn != DateTime.MinValue)
				{
                    //this.textBoxStart.Text = previewForm.MarkIn.ToString(Constants.dtformat);
                    this.textBoxStart.Text = Utility.GetLocalizedDateTimeString(previewForm.MarkIn, Constants.timeformat);
                }
				if (previewForm.MarkOut != DateTime.MinValue)
				{
                    //this.textBoxEnd.Text = previewForm.MarkOut.ToString(Constants.dtformat);
                    this.textBoxEnd.Text = Utility.GetLocalizedDateTimeString(previewForm.MarkOut, Constants.timeformat);
                }
			}

		}

		private void buttonDeckMatcher_Click(object sender, System.EventArgs e)
		{
			DeckMatcherForm deckMatcherForm = new DeckMatcherForm(((StreamGroup)this.listBoxPresentationSource.Items[0]).Cname,
				((StreamGroup)this.listBoxPresentationSource.Items[0]).Payload,
				this.textBoxStart.Text,this.textBoxEnd.Text,slideDecks);
			if (DialogResult.OK == deckMatcherForm.ShowDialog(this))
			{
				slideDecks = deckMatcherForm.ToSlideDecks();
			}
			
		}

		private void buttonRemovePresentation_Click(object sender, System.EventArgs e)
		{
			if (this.listBoxPresentationSource.Items.Count>0)
			{
				this.listBoxPresentationSource.Items.Clear();
				this.buttonRemovePresentation.Enabled = false;
				this.buttonDeckMatcher.Enabled = false;
                if (this.checkBoxSlidesReplaceVideo.Checked)
                {
                    this.checkBoxSlidesReplaceVideo.Checked = false;
                }
                this.checkBoxSlidesReplaceVideo.Enabled = false;
			}
		}

        private void checkBoxSlidesReplaceVideo_CheckedChanged(object sender, EventArgs e) {
            if (checkBoxSlidesReplaceVideo.Checked == true) {
                this.listBoxVideoSource.Items.Clear();
                this.listBoxVideoSource.Items.Add("[Presentation]");
                this.groupBoxVideo.Enabled = false;
                if (this.listBoxAudioSources.Items.Count > 0) {
                    this.buttonOk.Enabled = true;
                    this.buttonPreview.Enabled = true;
                }
            }
            else {
                this.listBoxVideoSource.Items.Clear();
                this.groupBoxVideo.Enabled = true;
                this.buttonPreview.Enabled = false;
                this.buttonOk.Enabled = false;
            }
        }

        /// <summary>
        /// Allow the user to choose a video source from which to build the presentation.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonPresentationFromVideo_Click(object sender, EventArgs e) {
            ChooseVideoForPresentationForm f = new ChooseVideoForPresentationForm(this.treeViewSources);
            DialogResult r = f.ShowDialog();
            if (r == DialogResult.OK) {
                if (this.checkBoxSlidesReplaceVideo.Checked) {
                    this.checkBoxSlidesReplaceVideo.Checked = false;
                }
                this.listBoxPresentationSource.Items.Clear();
                this.listBoxPresentationSource.Items.Add(f.ChosenVideoStreamGroup);
                this.buttonRemovePresentation.Enabled = true;
                this.buttonDeckMatcher.Enabled = false;
                this.checkBoxSlidesReplaceVideo.Enabled = false;

            }
        }

		#endregion UI Control Event Handlers

		#region Private Methods 

        /// <summary>
        /// CP3 uses a Guid for the cname.  Detect this case, and if possible, use the name to 
        /// make a more friendly participant name for display.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="cname"></param>
        /// <returns></returns>
        private string makeFriendlyCname(string name, string cname) {
            if ((name != null) && (name.Trim().Length > 0)) {
                try {
                    Guid g = new Guid(cname);
                    return name + " {" + cname + "}";
                }
                catch { }
            }
            return cname;
        }


        /// <summary>
        /// If a presentation source is specified, Check to see if decks exist and are matched.  If either is not
        /// the case, prompt the user to do so now.
        /// </summary>
        /// <returns></returns>
        private bool validateDecks(object sender, System.EventArgs e)
        {
            if (this.listBoxPresentationSource.Items.Count == 1)
            {
                if (slideDecks == null)
                {
                    StreamGroup sg = (StreamGroup)this.listBoxPresentationSource.Items[0];
                    if ((sg.Format == PresenterWireFormatType.CPNav) ||
                        (sg.Format == PresenterWireFormatType.CPCapability))
                    {
                        DialogResult dr = MessageBox.Show(this, "You have specified a presentation source, but no slide decks.  " +
                            "If no slide decks are specified, Archive Transcoder may not be able to build slide images." +
                            "\r\n\r\nDo you want to specify slide decks now?", "Do you want to specify slide decks?",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (dr == DialogResult.Yes)
                        {
                            this.buttonDeckMatcher_Click(sender, e);
                            return false;
                        }
                    }
                }
                else
                {
                    bool allmatched = true;
                    for (int i = 0; i < slideDecks.Length; i++)
                    {
                        if (!slideDecks[i].Matched)
                        {
                            allmatched = false;
                            break;
                        }
                    }
                    if (!allmatched)
                    {
                        DialogResult dr = MessageBox.Show(this, "The presentation data contains slide decks which " +
                            "have not been matched to files.  " +
                            "If the decks are not matched, Archive Transcoder may not be able to build slide images." +
                            "\r\n\r\nDo you want to match the slide decks now?", "Do you want to specify slide decks?",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (dr == DialogResult.Yes)
                        {
                            this.buttonDeckMatcher_Click(sender, e);
                            return false;
                        }
                    }
                }
            }
            return true;
        }

		/// <summary>
		/// Enforce that we have at least one audio, and either one video or presentation slides to replace video,
        /// and that start and end are valid as DateTimes, and that they define a positive duration timespan.
		/// </summary>
		/// <param name="start"></param>
		/// <param name="end"></param>
		/// <returns></returns>
		private bool validate(out DateTime start, out DateTime end)
		{
			start = DateTime.MinValue;
			end = DateTime.MinValue;

			if (this.listBoxAudioSources.Items.Count==0)
			{
				MessageBox.Show("You must select at least one audio source.");
				return false;
			}

            if ((this.listBoxVideoSource.Items.Count == 0) && (!this.checkBoxSlidesReplaceVideo.Checked))
            {
                MessageBox.Show("You must either select one video source, or select '" + this.checkBoxSlidesReplaceVideo.Text + "'");
                return false;
            }
			
			try
			{
				start = DateTime.Parse(this.textBoxStart.Text);
			}
			catch
			{
				MessageBox.Show("Failed to parse Start Time as a Date/Time: " + this.textBoxStart.Text);
				return false;
			}

			try
			{
				end = DateTime.Parse(this.textBoxEnd.Text);
			}
			catch
			{
				MessageBox.Show("Failed to parse End Time as a Date/Time: " + this.textBoxEnd.Text);
				return false;
			}

			if (start>=end)
			{
				MessageBox.Show("Time range is negative or zero.");
				return false;
			}
			return true;
		}

		private void setToolTips()
		{
			ToolTip tt = new ToolTip();
			tt.SetToolTip(this.buttonPreview,"Use preview to aid in locating segment start/end times.");
            tt.SetToolTip(this.checkBoxSlidesReplaceVideo, "Instead of the normal video stream, use the current slide image. This may be preferred for mobile devices.");
		}

		private String mkDetails(int streamCount, long duration, long startTime)
		{
			return mkDetails(streamCount, duration, startTime, PresenterRoleType.Other);
		}

		private String mkDetails(int streamCount, long duration, long startTime, PresenterRoleType role)
		{
			StringBuilder sb = new StringBuilder();
			sb.Append(" (");
			if (PresenterRoleType.Instructor == role)
			{
				sb.Append("Instructor; ");
			}
			if (streamCount > 1)
			{
				sb.Append(streamCount.ToString() + " streams; ");
			}
			if ((startTime != 0) && (startTime != long.MaxValue))
			{
				DateTime dt = new DateTime(startTime);
				sb.Append("start time=" + dt.ToString() + "; ");
			}
			sb.Append("duration=" + TimeSpan.FromSeconds(duration).ToString());
			sb.Append(")");
			return sb.ToString();
		}

		/// <summary>
		/// For the case where we are handed an existing segment to edit, we want to find out which conference
		/// it belongs to.  Assume the tree view has already been filled in from the database.  Just walk the tree until we find
		/// a match, or until we get to the end.  An alternative approach would be to store this information with 
		/// the segment, but we currently don't because I believe it is relevant only in this form.
		/// </summary>
		/// <param name="segment"></param>
		/// <returns></returns>
		private String GetCurrentConfString(ArchiveTranscoderJobSegment segment)
		{
			DateTime start = DateTime.MinValue;
			DateTime end = DateTime.MinValue;
			try
			{
				start = DateTime.Parse(segment.StartTime);
				end = DateTime.Parse(segment.EndTime);
			}
			catch
			{
				return "";
			}

			//consider it a match if the time is within range and at least one cname matches.
			String cname = "";
			if ((segment.VideoDescriptor !=null) && 
				(segment.VideoDescriptor.VideoCname != null) &&
				(segment.VideoDescriptor.VideoCname != ""))
			{
				cname = segment.VideoDescriptor.VideoCname;
			}
			else if ((segment.AudioDescriptor != null) && (segment.AudioDescriptor.Length > 0))
			{
				cname = segment.AudioDescriptor[0].AudioCname;
			}
			if ((cname==null) || (cname == ""))
				return "";

			foreach (TreeNode node in this.treeViewSources.Nodes)
			{
				if (node.Tag is Conference)
				{
					Conference c = (Conference)node.Tag;
					//Require that either the start or the end be within the range of the conference.
					if (((c.Start<=start) && (c.End>=start)) || 
						((c.Start<=end) && (c.End>=end)))
					{
						foreach (TreeNode pnode in node.Nodes)
						{
							if (pnode.Tag is Participant)
							{
								if (((Participant)(pnode.Tag)).CName == cname)
								{
									return c.Start.ToString() + " - " + c.Description;
								}
							}
						}
					}
				}
			}
			return "";
		}

		#endregion Private Methods

	}


}
