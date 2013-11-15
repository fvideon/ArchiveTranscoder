using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Sometimes the initial stream data lookup can be time consuming if there is a lot of new data.
	/// Draw a progress bar to show the user what's up.
	/// </summary>
	public class StreamLookupProgressForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.ProgressBar progressBar1;
		private System.Windows.Forms.Label labelStatus;
		private System.Windows.Forms.Button buttonCancel;
		private System.Windows.Forms.Label label2;
		private Thread waitThread;
		private bool stopNow;
		private ConferenceDataCache confDataCache;
		private System.ComponentModel.Container components = null;

		public StreamLookupProgressForm(ConferenceDataCache confDataCache)
		{
			InitializeComponent();
			this.confDataCache = confDataCache;

			this.labelStatus.Text = "Initializing ...";
			//start a new thread to update the progress bar, etc.
			waitThread = new Thread(new ThreadStart(waitThreadProc));
			waitThread.Name = "StreamLookupProgressForm Wait Thread";
			stopNow = false;
			this.progressBar1.Minimum = 0;
			this.progressBar1.Maximum = confDataCache.ParticipantCount;

			waitThread.Start();
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				stopThread();

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
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.labelStatus = new System.Windows.Forms.Label();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(16, 40);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(344, 23);
            this.progressBar1.TabIndex = 0;
            // 
            // labelStatus
            // 
            this.labelStatus.Location = new System.Drawing.Point(16, 80);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(248, 23);
            this.labelStatus.TabIndex = 1;
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(296, 80);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 2;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(24, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(344, 23);
            this.label2.TabIndex = 3;
            this.label2.Text = "Note: the speed of this process is related to the size of the database, and the a" +
                "mount of new data.";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // StreamLookupProgressForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(384, 126);
            this.ControlBox = false;
            this.Controls.Add(this.label2);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.labelStatus);
            this.Controls.Add(this.progressBar1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "StreamLookupProgressForm";
            this.ShowInTaskbar = false;
            this.Text = "Retrieving Stream Data From Database...";
            this.ResumeLayout(false);

		}
		#endregion

		private void stopThread()
		{
			stopNow = true;
			//if ((waitThread != null) && (waitThread.IsAlive))
			if ((waitThread != null))
			{
				if (!waitThread.Join(500))
				{
					Debug.WriteLine("Lookup wait thread aborting.");
					waitThread.Abort();
				}
			}
			waitThread = null;
		}

		private void buttonCancel_Click(object sender, System.EventArgs e)
		{
			stopThread();
			this.DialogResult = DialogResult.Cancel;
			this.Close();
		}

		private void waitThreadProc()
		{
			try
			{
				UpdateInvoke();
				bool done = false;
				int i = 0;
				while (!stopNow)
				{
					if(confDataCache.Working.WaitOne(100,false))
					{
						done=true;
						break;
					}
					//update progress bar and status.
					i++;
					if (i>=10)
					{
						UpdateInvoke();
						i=0;
					}
				}

				if ((done) && (!stopNow))
				{
					if (this.InvokeRequired)
					{
						this.Invoke(new WaitCompleteDelegate(WaitComplete));
					}
					else
					{
						WaitComplete();
					}
				}
			}
			catch (Exception ex)
			{
				//we catch at least the thread abort exception here from time to time
				Debug.WriteLine("waitThreadProc exiting with exception: " + ex.ToString());
			}
		}


		private void UpdateInvoke()
		{
			if (this.InvokeRequired)
			{
				this.Invoke(new UpdateDelegate(UpdateStatus));
			}
			else
			{
				UpdateStatus();
			}
		}

		private delegate void UpdateDelegate();
		private void UpdateStatus()
		{
			if (confDataCache.ParticipantCount >0)
			{
				if (confDataCache.ParticipantCount != this.progressBar1.Maximum)
					this.progressBar1.Maximum = confDataCache.ParticipantCount;

				this.progressBar1.Value = confDataCache.CurrentParticipant;
				this.labelStatus.Text = "Processing participant " + confDataCache.CurrentParticipant.ToString() + 
					" of " + confDataCache.ParticipantCount.ToString();
			}
			else
			{
				this.progressBar1.Value = 0;
				this.labelStatus.Text = "Connecting to Database ...";
			}
		}


		private delegate void WaitCompleteDelegate();
		private void WaitComplete()
		{
				this.DialogResult = DialogResult.OK;
				this.Close();
		}
	}
}
