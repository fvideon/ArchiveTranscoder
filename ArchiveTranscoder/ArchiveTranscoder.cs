using System;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Configuration;
using System.Text;
using System.Reflection;
using System.Collections;
using MSR.LST;
using MSR.LST.Net.Rtp;
using Microsoft.Win32;
using System.Collections.Generic;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Implement top level of UI-independent functionality.  Typically the app would create one instance 
	/// at launch and use it for the app lifetime.  There shouldn't be any problem creating multiple instances and using
	/// them simultaneously either, for instance the preview form creates a separate instance.
	/// </summary>
	public class ArchiveTranscoder:IDisposable
	{
		#region Private Members

		private ArchiveTranscoderBatch batch;	//The batch job currently loaded, or null
		private Thread workThread;				//Main transcoding thread.
		private bool endWorkThread;				//Flag that the thread should end asap.
		private WMSegment currentSegment;		//Job segment currently transcoding, or null
		private WMWriter wmWriter;				//The main Windows Media writer
		private String tempFileName;			//The output of the WM writer.
		private bool willOverwriteFiles;		//True if the currently loaded job will potentially cause files to be replaced
		private ProgressTracker progressTracker;//Keep tabs on the current transcoding job status
		private SlideImageGenerator slideImageGenerator; //Converts csd and ppt to image files
		private String sqlHost;					//Currently set ArchiveService database host
		private String dbName;					//Currently set database name
		private String batchLogFile;			//Log path for the last finished transcoder run
		private LogMgr batchLog;				//Maintains the log for the current (or most recent) transcoding job
		private bool pptInstalled;				//Indicates that PPT was or wasn't installed at construction time.
		private Hashtable rtDocuments;			//Collection of RTDocuments found in the presentation data.
        private Hashtable slideMessages;        //Collection of broadcast CP slides found in the presentation data.
        private Dictionary<Guid,string> frameGrabDirs;     //When video stills are used to make the presentation these temp directories contain the images.

		#endregion Private Members

		#region Public Properties

		/// <summary>
		/// The currently set ArchiveService database host
		/// </summary>
		public String SQLHost
		{
			get
			{
				return sqlHost;
			}
			set
			{
				sqlHost = value;
				resetConnectionString();
			}
		}

		/// <summary>
		/// The currently set Archive Service database name
		/// </summary>
		public String DBName
		{
			get
			{
				return dbName;
			}
			set
			{
				dbName = value;
				resetConnectionString();
			}
		}

		/// <summary>
		/// Indicates whether the currently loaded batch will potentially cause any files or
		/// directories to be overwritten.
		/// </summary>
		public bool WillOverwriteFiles
		{
			get {return willOverwriteFiles;}
		}

		/// <summary>
		/// The path to the log file for the most recently completed transcoder job.
		/// </summary>
		public String BatchLogFile
		{
			get {return batchLogFile;}
		}

		/// <summary>
		/// Value 0-10 indicating the most severe error encountered during the most recently
		/// completed transcoder job.  5 and up is worth warning the user, 7 and up is generally fatal.
		/// </summary>
		public int BatchErrorLevel
		{
			get 
			{
				if (batchLog != null)
					return batchLog.ErrorLevel;
				else
					return 0;
			}
		}

		/// <summary>
		/// The contents of the log as a string for the most recently completed transcoder job
		/// </summary>
		public String BatchLogString
		{
			get 
			{
				if (batchLog != null)
					return batchLog.ToString();
				else
					return "";
			}
		}

		#endregion Public Properties

		#region Construct/Dispose

		public ArchiveTranscoder()
		{
			batch = null;
			workThread = null;
			endWorkThread = false;
			currentSegment = null;
			wmWriter = null;
			tempFileName=null;
			willOverwriteFiles = false;
			progressTracker = null;
			slideImageGenerator = null;
			sqlHost=null;
			dbName=null;
			batchLogFile = null;
			batchLog = null;

			restoreRegSettings();
			resetConnectionString();
			pptInstalled = Utility.CheckPptIsInstalled();
		}

		public void Dispose()
		{
			saveRegSettings();

			if (progressTracker != null)
			{
				progressTracker.Stop();
				progressTracker = null;
			}
		}

		#endregion Construct/Dispose

		#region Public Methods

		/// <summary>
		/// Verify that the SqlHost is set and that we can successfully open a SqlConnection to that host.
		/// </summary>
		/// <returns></returns>
		public bool VerifyDBConnection()
		{
			if (this.sqlHost == "")
				return false;

			if (DatabaseUtility.CheckConnectivity())
				return true;

			return false;
		}

		/// <summary>
		/// Load the given batch.
		/// </summary>
		/// <param name="batch"></param>
		/// <returns>null if success, or error message string if there's any problem</returns>
		public String LoadJobs(ArchiveTranscoderBatch batch)
		{
			if (batch != null)
			{
				if ((this.workThread != null) &&
					(this.workThread.IsAlive))
				{					
					return "Cannot load job while encoding.";
				}
				else
				{
					this.batch = batch;
					willOverwriteFiles = this.outputFilesExist();
					return null;
				}
			}
			else
			{
				return "Null batch.";
			}
		}

		/// <summary>
		/// Start the loaded encoding job.  Return null if it starts correctly, otherwise return the error message.
		/// </summary>
		/// <returns>Error message if any.</returns>
		public String Encode()
		{
			if (batch==null)
			{
				return "Must set Encoding parameters first.";
			}

			if (workThread !=null)
			{
				if (workThread.IsAlive)
				{
					return "Already encoding.";
				}
			}
			
			String err = validateBatch(batch);
			if ((err != null) && (err.Trim() != ""))
			{
				return err;
			}

			//Start the encoding job in this thread:
			endWorkThread = false;
			workThread = new Thread(new ThreadStart(EncodingThread));
			workThread.Name = "Encoding Thread";
			//workThread.ApartmentState = ApartmentState.MTA; //I thought MTA was needed for WM interop, but maybe not??
            //In some circumstances we need a STA thread: exporting slide images from CP3 files that originated as XPS.
            workThread.SetApartmentState(ApartmentState.STA);
			workThread.Start();
			return null;
		}

		/// <summary>
		/// Cancel the job and terminate the active encoding thread, if any.
		/// </summary>
		public void Stop()
		{
			if (workThread == null)
				return;
			if (workThread.IsAlive)
			{
				endWorkThread=true;
				if (currentSegment != null)
					currentSegment.Stop();
				if (slideImageGenerator != null)
					slideImageGenerator.Stop();
				if (!workThread.Join(5000))
				{
					//This would be considered a bug if we have to abort.. it shouldn't ever happen.
					Debug.WriteLine("Aborting encoding thread.");
					workThread.Abort();
					if (File.Exists(tempFileName))
					{
						try
						{
							File.Delete(tempFileName);
						}
						catch (Exception e)
						{
							Debug.WriteLine("Failed to delete " + tempFileName + ". Exception text: " +
								e.ToString());
						}
					}
				}
				workThread = null;
			}
		}

		/// <summary>
		/// Launch a thread in which we stop the encoding thread, and raise an event when done.
		/// It's better to use this instead of Stop when the user clicks a button to stop.
		/// </summary>
		public bool AsyncStop()
		{
			Thread stopThread;
			if (workThread != null)
			{
				stopThread = new Thread(new ThreadStart(stopEncodingThread));
				stopThread.Name = "Stop Encoding Thread";
				stopThread.Start();
			}
			else
				return false;
			return true;
		}


		/// <summary>
		/// Examine current batch and return an error message if any segments have presentation data but no decks
		/// or if any decks are unmatched, or if a PPT deck is specified on a computer that doesn't have PPT installed,
		/// or if the deck is not a supported type.  
		/// These are non-fatal, but we want to warn/remind the user about this.
		/// </summary>
		/// <returns></returns>
		public String GetDeckWarnings()
		{
			StringBuilder errorMsg = new StringBuilder();
			if (batch==null)
			{
				return "";
			}
			foreach(ArchiveTranscoderJob job in this.batch.Job)
			{
				int segmentNum = 1;
				foreach (ArchiveTranscoderJobSegment segment in job.Segment)
				{
					if ((segment.PresentationDescriptor != null) && 
						(segment.PresentationDescriptor.PresentationCname != null) &&
						(segment.PresentationDescriptor.PresentationCname.Trim() != ""))
					{
						if (segment.SlideDecks == null)
						{
							if ((segment.PresentationDescriptor.PresentationFormat == "CPNav") ||
								(segment.PresentationDescriptor.PresentationFormat == "CPCapability"))
							{
								errorMsg.Append("Job " + job.ArchiveName + ", segment " + segmentNum.ToString() +
									" includes a presentation source, but no slide decks are specified. \r\n");
							}
						}
						else
						{
							foreach (ArchiveTranscoderJobSlideDeck deck in segment.SlideDecks)
							{
								if ((deck.Path == null) || (deck.Path.Trim() == ""))
								{
									errorMsg.Append("Job " + job.ArchiveName + ", segment " + segmentNum.ToString() +
										" includes an unmatched slide deck: " + deck.Title + ". \r\n");
								}
								else if ((!pptInstalled) && ((deck.Path.ToLower().EndsWith("ppt")) || (deck.Path.ToLower().EndsWith("pptx"))))
								{
									errorMsg.Append("Job " + job.ArchiveName + ", segment " + segmentNum.ToString() +
										" specifies a PowerPoint deck, but PowerPoint does not appear to be installed on this computer: " + 
										deck.Path + ". \r\n");
								}
								else if ((!deck.Path.ToLower().EndsWith("ppt")) && 
                                    (!deck.Path.ToLower().EndsWith("csd")) &&
                                    (!deck.Path.ToLower().EndsWith("cp3")) &&
                                    (!deck.Path.ToLower().EndsWith("pptx")))
								{
                                    if (Directory.Exists(deck.Path)) { 
                                        //It might be a directory containing image files.. could validate further here.
                                    }
                                    else {
                                        errorMsg.Append("Job " + job.ArchiveName + ", segment " + segmentNum.ToString() +
                                            " specifies a slide deck which does not appear to be a supported type: " +
                                            deck.Path + ".  PPT, PPTX, CP3 and CSD deck types are supported.\r\n");
                                    }
								}
							}
						}
					}
					segmentNum++;
				}
			}
			return errorMsg.ToString();
		}

		#endregion Public Methods

		#region Private Methods

		#region Batch Validation

		/// <summary>
		/// Validate the given batch.  Return "" if all is ok, otherwise return a list of errors.
		/// </summary>
		/// <param name="batch"></param>
		/// <returns></returns>
		private String validateBatch(ArchiveTranscoderBatch batch)
		{
			if (batch == null)
				return "Batch is null.";

			StringBuilder err = new StringBuilder();
			if ((batch.Job==null) || (batch.Job.Length==0))
			{
				err.Append("Batch contains no jobs. \r\n");
			}

			//It is ok if server and dbname are not specified.  We will use the defaults.

			foreach (ArchiveTranscoderJob job in batch.Job)
			{
				if (job==null)
					err.Append("Null job found.\r\n");
				else
					err.Append(validateJob(job));
			}

			return err.ToString();
		}

		private String validateJob(ArchiveTranscoderJob job)
		{
			StringBuilder err = new StringBuilder();

			if (!Utility.isSet(job.ArchiveName))
			{
				err.Append("Job ArchiveName is not set. \r\n");
			}

			bool validOutDir = true;

			if (!Utility.isSet(job.Path))
			{
				validOutDir = false;
				err.Append("Job Output Path is not set. \r\n");
			}
			else
			{
				try
				{
					string s = Path.GetFullPath(job.Path);
				}
				catch
				{
					validOutDir = false;
					err.Append("Job Output Path is invalid. \r\n");
				}
			}

			if (!Utility.isSet(job.BaseName))
			{
				err.Append("Job BaseName is not set. \r\n");
			}
			else if (Utility.ContainsSlash(job.BaseName)) 
			{
				err.Append("Job BaseName contains invalid characters. \r\n");			
			}
			else if (validOutDir)
			{
				try
				{
					string s = Path.GetFullPath(Path.Combine(job.Path,job.BaseName));
				}
				catch
				{
					validOutDir = false;
					err.Append("Job BaseName contains invalid characters. \r\n");
				}
			}

            bool norecompression = false;
			if (!Utility.isSet(job.WMProfile))
			{
				err.Append("Job Windows Media Profile is not set. \r\n");
			}
            else
            {
                if (job.WMProfile.ToLower() == "norecompression")
			    {
                    norecompression = true;
                }
            }
			
			if ((job.Segment == null) || (job.Segment.Length==0))
			{
				err.Append("Job contains no segments. \r\n");
			}
			else
			{
				int segmentNum = 1;
                bool useSlideStream = false;
				foreach (ArchiveTranscoderJobSegment segment in job.Segment)
				{
                    if (segment == null)
                    {
                        err.Append("Segment " + segmentNum.ToString() + " is null \r\n");
                    }
                    else
                    {
					    err.Append(validateSegment(segment, segmentNum));

                        if (Utility.SegmentFlagIsSet(segment, SegmentFlags.SlidesReplaceVideo))
                        {
                            useSlideStream = true;
                        }
                    }
					segmentNum++;
				}
                if (useSlideStream && norecompression)
                {
                    err.Append("No recompression mode is not supported for jobs using slides in place of video.  Specify a Windows Media profile for this job. \r\n");
                }
			}

			if ((job.Target == null) || (job.Target.Length==0))
			{
				err.Append("Job contains no targets. \r\n");
			}
			else
			{
				foreach (ArchiveTranscoderJobTarget target in job.Target)
				{
					err.Append(validateTarget(target));
				}
			}

			return err.ToString();
		}

		private String validateSegment(ArchiveTranscoderJobSegment segment, int segmentNum)
		{
			StringBuilder err = new StringBuilder();
			if (!Utility.isDateTime(segment.StartTime))
				err.Append("Could not parse segment " + segmentNum.ToString() + " start time \r\n");
			if (!Utility.isDateTime(segment.EndTime))
				err.Append("Could not parse segment " + segmentNum.ToString() + " end time \r\n");
			
			//Require that at least one audio source
			if ((segment.AudioDescriptor == null) || (segment.AudioDescriptor.Length==0))
			{
				err.Append("Segment " + segmentNum.ToString() + " must specifiy at least one audio source. \r\n");
			}
			else
			{
				foreach (ArchiveTranscoderJobSegmentAudioDescriptor ad in segment.AudioDescriptor)
				{
					if (!isSet(ad))
					{
						err.Append("Segment " + segmentNum.ToString() + " contains a null audio source. \r\n");
					}
				}
			}

            //if the "slides replace video" flag is set, there must be a presentation source.  Otherwise there must be a video source.
            if (Utility.SegmentFlagIsSet(segment, SegmentFlags.SlidesReplaceVideo))
            {
                if (segment.PresentationDescriptor == null)
                {
                    err.Append("Segment " + segmentNum.ToString() + ": If slides are used in place of video, there must be a presentation source specified. \r\n");
                }
                //later we verify that there is at least a cname.
            }
            else
            {
                if (!isSet(segment.VideoDescriptor))
                {
                    err.Append("Segment " + segmentNum.ToString() + " must specifiy at least one video source. \r\n");
                }
            }

			if (segment.PresentationDescriptor != null)
			{
				if (!Utility.isSet(segment.PresentationDescriptor.PresentationCname))
				{
					err.Append("Segment " + segmentNum.ToString() + " contains a presentation descriptor with an unspecified source. \r\n");
				}
			}

			if (segment.SlideDecks != null)
			{
				foreach (ArchiveTranscoderJobSlideDeck deck in segment.SlideDecks)
				{
					err.Append(validateSlideDeck(deck, segmentNum));
				}
			}
			return err.ToString();
		
		}

		private String validateSlideDeck(ArchiveTranscoderJobSlideDeck deck, int segmentNum)
		{
			StringBuilder err = new StringBuilder();
			
			if (!Utility.isSet(deck.Path))
			{
				//Unmatched deck: the user has been warned about this, so do nothing here:
				//err.Append("Slide deck path not specified in segment " + segmentNum.ToString() + ".\r\n");
			}
			else
			{
				if (!File.Exists(deck.Path))
				{
                    if (!Directory.Exists(deck.Path)) {
                        err.Append("Could not find slide deck " + deck.Path + " in segment " + segmentNum.ToString() + ".\r\n");
                    }
				}
			}

			try
			{
				Guid g = new Guid(deck.DeckGuid);
			}
			catch
			{
				err.Append("Could not parse slide deck guid as Guid: " + deck.DeckGuid + " in segment " + segmentNum.ToString() + ".\r\n");
			}

			return err.ToString();
		}


		private String validateTarget(ArchiveTranscoderJobTarget target)
		{
			StringBuilder err = new StringBuilder();
			if (target.Type == null)
			{
				err.Append("Target type not specified. \r\n");
			}
			else
			{
				if (target.Type.ToLower().Trim()=="stream")
				{
					if (target.SlideBaseUrl!=null)
					{
						if (!Utility.isUri(target.SlideBaseUrl))
						{
							err.Append("Slide Base URL cannot be parsed as a URI. \r\n");
						}
					}
					if (target.CreateAsx != null)
					{
						if (target.CreateAsx.Trim().ToLower()=="true")
						{
							if (target.WmvUrl == null)
							{
								err.Append("To create ASX file, WMV URL must be specified. \r\n");
							} 
							else
							{
								if (!Utility.isUri(target.WmvUrl))
								{
									err.Append("Could not parse WMV URL. \r\n");
								}
							}
							if (target.PresentationUrl != null)
							{
								if (!Utility.isUri(target.PresentationUrl))
								{
									err.Append("Could not parse Presentation URL. \r\n");
								}
							}
						}
					}
					if (target.CreateWbv != null)
					{
						if (target.CreateWbv.Trim().ToLower()=="true")
						{
							if (target.AsxUrl==null)
							{
								err.Append("To create WBV file, ASX URL must be specified. \r\n");
							}
							else
							{
								if (!Utility.isUri(target.AsxUrl))
								{
									err.Append("Could not parse ASX URL. \r\n");
								}
							}

							if (target.CreateAsx.Trim().ToLower()!="true")
							{
								err.Append("To create a WBV file, you must create a ASX file also. \r\n");
							}
						}
					}
				}
				else if (target.Type.ToLower().Trim()=="download")
				{
					//no additional params?
				}
                else {
                    err.Append("Unrecognized target type: " + target.Type + "\r\n");
                }
			}
			return err.ToString();
			
		}



		private bool isSet(ArchiveTranscoderJobSegmentVideoDescriptor videoDescriptor)
		{
			if (videoDescriptor != null)
			{
				return Utility.isSet(videoDescriptor.VideoCname);
			}
			return false;
		}

		private bool isSet(ArchiveTranscoderJobSegmentAudioDescriptor audioDescriptor)
		{
			if (audioDescriptor != null)
			{
				return Utility.isSet(audioDescriptor.AudioCname);
			}
			return false;
		}

		/// <summary>
		/// Make sure there is some data for each stream in each segment in the job.
		/// </summary>
		/// <param name="job"></param>
		/// <param name="jobIndex"></param>
		/// <returns></returns>
		private bool validateStreams(ArchiveTranscoderJob job, int jobIndex, LogMgr log)
		{
			String emptyStreams;
			bool errorFound = false;
			for (int i=0; i<job.Segment.Length;i++)
			{
				ArchiveTranscoderJobSegment segment = job.Segment[i];
				emptyStreams = findEmptyStreams(segment);
				if (emptyStreams != null)
				{
					//jobErrorMessages[jobIndex] += "Empty streams found in segment " + (i+1).ToString() + ": " +
					//	emptyStreams + ".  ";
					errorFound = true;
					log.WriteLine("Error: Empty streams found in segment " + (i+1).ToString() + ": " +
						emptyStreams + ".  ");
					log.ErrorLevel=7;
				}
			}
			if (errorFound)
			{
				return false;
			}
			return true;
		}


		/// <summary>
		/// Make sure that all of the audio and video streams in segment contain at least some data.
        /// If "slides replace video" flag is set, make sure there is some presentation data.
		/// Return a list of any that are empty, or null if all looks good.
		/// </summary>
		/// <param name="segment"></param>
		/// <returns></returns>
		private String findEmptyStreams(ArchiveTranscoderJobSegment segment)
		{
			String emptyStreams = "";
			long start = DateTime.Parse(segment.StartTime).Ticks;
			long end = DateTime.Parse(segment.EndTime).Ticks;


			for(int i=0; i< segment.AudioDescriptor.Length; i++)
			{
				ArchiveTranscoderJobSegmentAudioDescriptor ad = segment.AudioDescriptor[i];
				if (!Utility.isSet(ad.AudioName))
				{
					if (DatabaseUtility.isEmptyStream(PayloadType.dynamicAudio,ad.AudioCname,null,start,end))
					{
						emptyStreams += "Audio:" + ad.AudioCname;
					}
				}
				else
				{
					if (DatabaseUtility.isEmptyStream(PayloadType.dynamicAudio,ad.AudioCname,
						ad.AudioName,start,end))
					{
						emptyStreams += "Audio:" + ad.AudioCname + " - " +
							ad.AudioName;
					}
				}
			}

            //If "slides replace video" flag is set, we require a presentation source, otherwise require a video source.
            if (Utility.SegmentFlagIsSet(segment, SegmentFlags.SlidesReplaceVideo))
            {
                if (DatabaseUtility.isEmptyStream(PayloadType.dynamicPresentation, segment.PresentationDescriptor.PresentationCname, null, start, end))
                {
                    if (DatabaseUtility.isEmptyStream(PayloadType.RTDocument, segment.PresentationDescriptor.PresentationCname, null, start, end))
                    {
                        //We don't consider this case to be an error.
                        //emptyStreams += "Presentation (slides to replace video):" + segment.PresentationDescriptor.PresentationCname;
                    }
                }
            }
            else
            {
                if (!Utility.isSet(segment.VideoDescriptor.VideoName))
                {
                    if (DatabaseUtility.isEmptyStream(PayloadType.dynamicVideo, segment.VideoDescriptor.VideoCname, null, start, end))
                    {
                        emptyStreams += "Video:" + segment.VideoDescriptor.VideoCname;
                    }
                }
                else
                {
                    if (DatabaseUtility.isEmptyStream(PayloadType.dynamicVideo, segment.VideoDescriptor.VideoCname,
                        segment.VideoDescriptor.VideoName, start, end))
                    {
                        emptyStreams += "Video:" + segment.VideoDescriptor.VideoCname + " - " +
                            segment.VideoDescriptor.VideoName;
                    }
                }
            }

			if (emptyStreams == "")
				return null;
			return emptyStreams;
		}

		#endregion Batch Validation

		#region Transcoding Thread Proc

		private void EncodingThread()
		{
            try {
                PresentationStreamWriter pStreamWriter = null;
                ArchiveTranscoderJob job;
                WMSegment prevSegment;
                String errMsg = null;
                ProfileData profileData = null;
                String prxFile = "";
                uint profileIndex = 0;
                bool noRecompression = false;
                rtDocuments = null;

                //This log is written to the output as part of packaging.
                LogMgr jobLog;

                //This log is an accumulation of all logs for the batch which is made available to the caller.
                batchLogFile = null;
                batchLog = new LogMgr();
                batchLog.WriteLine("ArchiveTranscoder starting.");

                tempFileName = "";
                for (int i = 0; i < this.batch.Job.Length; i++) {
                    jobLog = new LogMgr();
                    jobLog.WriteLine("Processing job " + (i + 1).ToString() + " of " + this.batch.Job.Length.ToString() + ".");

                    job = batch.Job[i];

                    if (progressTracker != null) {
                        progressTracker.Stop();
                    }
                    progressTracker = new ProgressTracker(job.Segment.Length);
                    progressTracker.OnShowProgress += new ProgressTracker.showProgressCallback(OnShowProgress);
                    progressTracker.Start();
                    progressTracker.AVStatusMessage = "Initializing";

                    /// Originally we had code here to enforce that all segments had at least some data
                    /// but that constraint was relaxed I think to work around the case of a presentation stream
                    /// in a very short segment that had no data.  Now we skip segments or jobs below only where
                    /// necessary.

                    if (endWorkThread) {
                        progressTracker.Stop();
                        progressTracker = null;
                        break;
                    }

                    if ((job.WMProfile != null) && (job.WMProfile.ToLower() == "norecompression")) {
                        // a special case "native" profile will write a WMV with one audio and one video without recompression.
                        //PRI2: There is a bug here:  If the job contains multiple audio streams we will make the profile
                        // arbitrarily using the first one in the first segment.  When the audio streams are mixed, a
                        // different native profile may be chosen.
                        //profileString = ProfileUtility.MakeNativeProfile(job.Segment[0]);
                        //profileData = ProfileUtility.SegmentToProfileData(job.Segment[0]);
                        profileData = jobToProfileData(job);
                        if (profileData == null) {
                            //This is null if there are no streams that contain some data
                            progressTracker.AVStatusMessage = "No valid streams found.";
                            progressTracker.Stop();
                            progressTracker = null;
                            jobLog.WriteLine("Skipping job because empty streams were found.");
                            jobLog.ErrorLevel = 7;
                            batchLog.Append(jobLog);
                            continue;
                        }
                        noRecompression = true;
                    }
                    else {
                        //First try to parse the profile as an int.  If that fails, assume it's a file.
                        if (!uint.TryParse(job.WMProfile, out profileIndex)) {
                            prxFile = job.WMProfile;
                            if (!File.Exists(prxFile)) {
                                //If the prx file isn't found try looking in the app directory.
                                //This is to hack around an issue with working between x64 and x86 systems with different Program Files paths.
                                jobLog.WriteLine("Prx file doesn't exist: " + prxFile);
                                System.Reflection.Assembly a = System.Reflection.Assembly.GetExecutingAssembly();
                                String appDir = System.IO.Path.GetDirectoryName(a.Location);
                                prxFile = Path.Combine(appDir, Path.GetFileName(prxFile));
                                if (File.Exists(prxFile)) {
                                    jobLog.WriteLine("Using prx file: " + prxFile);
                                }
                                else {
                                    jobLog.WriteLine("Failed to locate suitable prx file.");
                                }
                            }
                        }


                        noRecompression = false;
                    }

                    if (endWorkThread) {
                        progressTracker.Stop();
                        progressTracker = null;
                        jobLog.WriteLine("Encoding ended by user.");
                        batchLog.Append(jobLog);
                        break;
                    }

                    wmWriter = new WMWriter();
                    wmWriter.Init();
                    bool result = false;
                    if (noRecompression) {
                        result = wmWriter.ConfigProfile(profileData);
                    }
                    else {
                        result = wmWriter.ConfigProfile(null, prxFile, profileIndex);
                    }

                    if (!result) {
                        jobLog.WriteLine("Failed to configure Windows Media Profile.");
                        jobLog.ErrorLevel = 7;
                        progressTracker.AVStatusMessage = "Failed to configure profile.";
                        progressTracker.Stop();
                        progressTracker = null;
                        batchLog.Append(jobLog);
                        continue;
                    }

                    //PRI3: estimate disk space requirements and verify that the current volume has enough space.
                    tempFileName = Utility.GetTempFilePath("wmv");
                    wmWriter.ConfigFile(tempFileName);

                    if (noRecompression) {
                        wmWriter.ConfigNullProps();
                        //Note: it's not clear that this really matters:
                        //wmWriter.SetCodecInfo();
                    }

                    //Warning: test code here:
                    //wmWriter.SetFixedFrameRate(true);

                    wmWriter.Start();


                    if (containsPresentationData(job)) {
                        pStreamWriter = new PresentationStreamWriter(jobLog);
                    }
                    else {
                        pStreamWriter = null;
                    }

                    prevSegment = null;
                    long jobStart = DateTime.Parse(job.Segment[0].StartTime).Ticks; //The absolute time of the start of the job
                    long offset = 0; //The relative AV time to write to to begin the next segment in ticks
                    List<WMSegment> segments = new List<WMSegment>();
                    for (int j = 0; j < job.Segment.Length; j++) {
                        jobLog.WriteLine("Processing segment " + (j + 1).ToString() + " of " + job.Segment.Length.ToString() + ".");
                        progressTracker.CurrentSegment = j + 1;
                        string emptyStreams = findEmptyStreams(job.Segment[j]);
                        if (emptyStreams != null) {
                            jobLog.WriteLine("Skipping segment because empty streams were found.");
                            continue;
                        }

                        currentSegment = new WMSegment(job, job.Segment[j], jobStart, offset, wmWriter, progressTracker, jobLog, profileData, noRecompression, prevSegment);
                        segments.Add(currentSegment);

                        if (endWorkThread)
                            break;

                        errMsg = currentSegment.Encode();

                        if (errMsg != null) {
                            jobLog.WriteLine("Error: " + errMsg);
                            jobLog.ErrorLevel = 7;
                            break;
                        }

                        if (endWorkThread) {
                            break;
                        }

                        offset += currentSegment.ActualWriteDuration + 1;
                        prevSegment = currentSegment;
                    }

                    //Close WMWriter
                    if (wmWriter != null) {
                        wmWriter.Stop();
                        wmWriter.Cleanup();
                        wmWriter = null;
                    }

                    if ((errMsg == null) && (!endWorkThread)) {
                        //If there are frame grabs, filter them as needed.
                        PresentationFromVideoMgr.FilterSegments(segments);
                        int j = 0;
                        foreach (WMSegment s in segments) {
                            if (s.ContainsPresentationData) {
                                if (s.DeckTitles != null) {
                                    logMissingDeckWarnings(jobLog, job, s.DeckTitles, j + 1);
                                }
                                pStreamWriter.Write(s.GetPresentationData());
                                pStreamWriter.Flush();
                                //For RTDocuments, the currentSegment may have slide images as well.
                                // We need to keep a reference to them here.
                                accumulateRTDocuments(s.RTDocuments, jobLog);
                                accumulateSlideMessages(s.SlideMessages, jobLog);
                                accumulateVideoStillsDirectory(s);
                            }
                            j++;
                        }
                    }

                    //Close Presentation StreamWriter
                    if (pStreamWriter != null) {
                        pStreamWriter.Close();
                    }

                    //Make slide images, and package output
                    if ((errMsg == null) && (!endWorkThread)) {
                        if (pStreamWriter != null) {
                            slideImageGenerator = SlideImageGenerator.GetInstance(job, progressTracker, jobLog);
                            slideImageGenerator.TheRtDocuments = this.rtDocuments;
                            slideImageGenerator.TheSlideMessages = this.slideMessages;
                            slideImageGenerator.SetImageExportSize(true, 0, 0);
                            slideImageGenerator.Process();
                        }
                        if (!endWorkThread)
                            errMsg = packageJob(job, tempFileName, pStreamWriter, slideImageGenerator, jobLog);
                    }

                    if ((errMsg != null) || (endWorkThread)) {
                        //cleanup and exit thread.
                        if (File.Exists(tempFileName))
                            File.Delete(tempFileName);
                        progressTracker.Stop();
                        progressTracker = null;
                        slideImageGenerator = null;
                        if (endWorkThread) {
                            jobLog.WriteLine("Encoding ended by user.");
                        }
                        batchLog.Append(jobLog);
                        break;
                    }
                    progressTracker.Stop();
                    progressTracker = null;
                    slideImageGenerator = null;
                    batchLog.Append(jobLog);
                }//.. and on to the next job in the batch ..

                batchLogFile = batchLog.WriteTempFile();

                pStreamWriter = null;
                if (OnBatchCompleted != null) {
                    OnBatchCompleted();
                }
            }
            catch (Exception ex) {
                Console.WriteLine(ex.ToString());
                throw;
            }
		}


 		#endregion Transcoding Thread Proc

		#region Output Packaging

		private String packageJob(ArchiveTranscoderJob job, String wmvFile, PresentationStreamWriter pStreamWriter,
			SlideImageGenerator slideImageGenerator, LogMgr jobLog)
		{
			if (endWorkThread)
				return null;

			String errMsg  = null;
			progressTracker.CustomMessage = "Packaging output.";
			progressTracker.ShowPercentComplete = false;

			if (!Directory.Exists(job.Path))
			{
				Directory.CreateDirectory(job.Path);
			}
			String streamDir = Path.Combine(job.Path,job.BaseName+@"\stream\");
			String downloadDir = Path.Combine(job.Path,job.BaseName+@"\download\");

			foreach (ArchiveTranscoderJobTarget t in job.Target)
			{
				if (t.Type.ToLower().Trim() == "stream")
				{
					try
					{
						errMsg = packageStream(streamDir, job, t, wmvFile, pStreamWriter, slideImageGenerator,jobLog);
					}
					catch (Exception e)
					{
						jobLog.WriteLine("Failure while packaging stream target: " + e.ToString());
						jobLog.ErrorLevel = 5;
					}
				}
				else if (t.Type.ToLower().Trim() == "download")
				{
					try
					{
						errMsg = packageDownload(downloadDir, job, t, wmvFile, pStreamWriter, slideImageGenerator,jobLog);
					}
					catch (Exception e)
					{
						jobLog.WriteLine("Failure while packaging download target: " + e.Message);
						jobLog.ErrorLevel = 5;
					}
				}
				else
				{
					Debug.WriteLine("Unhandled target type.");
				}

				if (endWorkThread)
					break;
			}

			String logTempPath = jobLog.WriteTempFile();
			String logPath = Path.Combine(job.Path,job.BaseName + @"\log.txt");
			File.Copy(logTempPath,logPath,true);

			if (File.Exists(wmvFile))
			{
				File.Delete(wmvFile);
			}

            if (this.frameGrabDirs != null) {
                foreach (string d in this.frameGrabDirs.Values) {
                    Directory.Delete(d, true);
                }
                this.frameGrabDirs = null;
            }

			return errMsg;
		}

		/// <summary>
		/// Do packaging for streaming target.
		/// </summary>
		/// <param name="outDir"></param>
		/// <param name="job"></param>
		/// <param name="wmvFile"></param>
		/// <param name="presFile"></param>
		/// <returns></returns>
		private String packageStream(String outDir, ArchiveTranscoderJob job, 
			ArchiveTranscoderJobTarget target, String wmvFile,  PresentationStreamWriter pStreamWriter,
			SlideImageGenerator slideImageGenerator,LogMgr jobLog)
		{
			progressTracker.CustomMessage = "Packaging Stream Target.";
			jobLog.WriteLine("Packaging Stream Target.");
			StringBuilder deployNotes = new StringBuilder();
			deployNotes.Append("Deployment Notes \r\n \r\n");
			deployNotes.Append("This directory contains the files necessary for a streamed archive.  " +
				"The following tasks must be completed in order to deploy the archive. \r\n");

			if (!Directory.Exists(outDir))
			{
				Directory.CreateDirectory(outDir);
			}
						
			//Set the local output names for presentation and WMV
			String wmvBaseName = getBaseFileName(target.WmvUrl,job.BaseName) + ".wmv";
			String destFile = outDir + wmvBaseName;

			String presBaseName = getBaseFileName(target.PresentationUrl,job.BaseName) + ".xml";
			String pFile = outDir +  presBaseName;

			//Construct the deployment URLs for presentation and WMV
			String WmvUrl = null;
			if (target.WmvUrl != null)
			{
				if (target.WmvUrl.Trim().ToLower().EndsWith("wmv"))
				{
					WmvUrl = target.WmvUrl.Trim();
				}
				else
				{
					String wmvFileName = getBaseFileName(target.WmvUrl,job.BaseName) + ".wmv";
					WmvUrl = combineUrl(target.WmvUrl,wmvFileName);	
				}

				deployNotes.Append("-Make the WMV file " + wmvBaseName + " accessible from " + WmvUrl + ". \r\n");
			}
			else
			{
				deployNotes.Append("-Make the WMV file " + wmvBaseName + " available from your Windows Media Server or webserver. \r\n");
			}

			String presUrl = null;
			if ((pStreamWriter != null) && (target.PresentationUrl != null))
			{
				if (target.PresentationUrl.Trim().ToLower().EndsWith(".xml"))
				{
					presUrl = target.PresentationUrl.Trim();
				}
				else
				{
					String presFileName = getBaseFileName(target.PresentationUrl,job.BaseName) + ".xml";
					presUrl = combineUrl(target.PresentationUrl,presFileName);
				}
				deployNotes.Append("-Make the presentation data file " + presBaseName + " available from " + presUrl + ". \r\n");
			}

			//copy WMV to its local target
			if (File.Exists(destFile))
			{
				try
				{
					File.Delete(destFile);
				}
				catch (Exception e)
				{
					jobLog.WriteLine("Failed to copy WMV file to " + destFile + ".  " + e.Message);
					jobLog.ErrorLevel=7;
					return null;
				}
			}
			//PRI3: if the user clicks Stop during this copy we have no recourse but to abort the thread, right?
			File.Copy(wmvFile,destFile);
			//copy presentation data and slides to the local target.
			if (pStreamWriter != null)
			{
				String presFile = pStreamWriter.WriteHeaderAndCopy(job.Segment[0].StartTime,target.SlideBaseUrl,"jpg",job.ArchiveName);
				if (presFile != null)
				{
					if (File.Exists(pFile))
					{
						File.Delete(pFile);
					}
					File.Move(presFile,pFile);
				}
				if (slideImageGenerator != null)
				{
					foreach (Guid g in slideImageGenerator.OutputDirs.Keys)
					{
						DirectoryInfo diOutDir = new DirectoryInfo(Path.Combine(outDir,g.ToString()));
						if (diOutDir.Exists)
						{
							try 
							{
								diOutDir.Delete(true);
							}
							catch (Exception e)
							{
								jobLog.WriteLine("Failed to write output directory: " + diOutDir.FullName + ".  " + e.Message);
								jobLog.ErrorLevel = 7;
								return null;
							}
						}
						diOutDir.Create();
						DirectoryInfo diSrcDir = new DirectoryInfo((String)slideImageGenerator.OutputDirs[g]);
						FileInfo[] files = diSrcDir.GetFiles();
						foreach (FileInfo f in files)
						{
							f.CopyTo(Path.Combine(diOutDir.FullName,f.Name));
						}
						if (target.SlideBaseUrl != null)
						{
							deployNotes.Append("-Make the directory of slide images " + diOutDir.Name + " available from " +
								combineUrl(target.SlideBaseUrl,diOutDir.Name) + ". \r\n");
						}
					}
				}
                if (this.frameGrabDirs != null) {
                    foreach (Guid g in this.frameGrabDirs.Keys) {
                        DirectoryInfo diOutDir = new DirectoryInfo(Path.Combine(outDir, g.ToString()));
                        if (diOutDir.Exists) {
                            try {
                                diOutDir.Delete(true);
                            }
                            catch (Exception e) {
                                jobLog.WriteLine("Failed to write output directory: " + diOutDir.FullName + ".  " + e.Message);
                                jobLog.ErrorLevel = 7;
                                return null;
                            }
                        }
                        diOutDir.Create();
                        DirectoryInfo diSrcDir = new DirectoryInfo(this.frameGrabDirs[g]);
                        FileInfo[] files = diSrcDir.GetFiles();
                        foreach (FileInfo f in files) {
                            f.CopyTo(Path.Combine(diOutDir.FullName, f.Name));
                        }
                        if (target.SlideBaseUrl != null) {
                            deployNotes.Append("-Make the directory of slide images " + diOutDir.Name + " available from " +
                                combineUrl(target.SlideBaseUrl, diOutDir.Name) + ". \r\n");
                        }
                    }
                }
			}

			//Create the asx and wbv
			if (Convert.ToBoolean(target.CreateAsx))
			{
				String asxFileName = getBaseFileName(target.AsxUrl,job.BaseName) + ".asx";
				String asxPath = Path.Combine(outDir,asxFileName);
				//set asx deployment location
				String asxUrl = "";
				if (target.AsxUrl.Trim().ToLower().EndsWith(".asx"))
					asxUrl = target.AsxUrl.Trim();
				else 
					asxUrl = combineUrl(target.AsxUrl, asxFileName);

				writeAsx(asxPath,WmvUrl,presUrl,job.ArchiveName); 


				if (Convert.ToBoolean(target.CreateWbv))
				{
					deployNotes.Append("-Make the ASX file " + asxFileName + " available from " + asxUrl + ". \r\n" );
					writeWbv(Path.Combine(outDir,job.BaseName+".wbv"), asxUrl);
					deployNotes.Append("-Make the WBV file " + job.BaseName + ".wbv" + " available to your WebViewer clients " +
						"by posting it on your website, or email, etc. \r\n");
				}
				else
				{
					deployNotes.Append("-Make the ASX file " + asxFileName + 
						" available to your clients by posting it on your website, etc. \r\n");
				}
			}

			StreamWriter deploySw = File.CreateText(Path.Combine(outDir,"Deployment Notes.txt"));
			deploySw.Write(deployNotes.ToString());
			deploySw.Close();
			return null;
		}

		/// <summary>
		/// Prepare download package.
		/// </summary>
		/// <param name="outDir"></param>
		/// <param name="job"></param>
		/// <param name="wmvFile"></param>
		/// <param name="presFile"></param>
		/// <returns></returns>
		private String packageDownload(String outDir, ArchiveTranscoderJob job, 
			ArchiveTranscoderJobTarget target, String wmvFile,  PresentationStreamWriter pStreamWriter,
			SlideImageGenerator slideImageGenerator, LogMgr jobLog)
		{
			progressTracker.CustomMessage = "Packaging Download Target.";
			jobLog.WriteLine("Packaging Download Target.");
			StringBuilder deployNotes = new StringBuilder();

			//outDir is the directory named <basename>\download.  Create it if it doesn't exist.
			if (!Directory.Exists(outDir))
			{
				DirectoryInfo diOutDir = Directory.CreateDirectory(outDir);
			}

			//Create a subdirectory of outdir in which to collect relevant files.  We will zip and delete this
			// directory when done.
			String jobdirname = Path.Combine(outDir,job.BaseName);
			//This one we will delete if it does exist, and create anew
			if (Directory.Exists(jobdirname))
			{
				try
				{
					Directory.Delete(jobdirname,true);
				}
				catch (Exception e)
				{
					jobLog.WriteLine("Warning: Failed to create output directory: " + jobdirname + "  " + e.Message);
					jobLog.ErrorLevel = 7;
					return null;
				}
			}

			DirectoryInfo jobdir = Directory.CreateDirectory(jobdirname);
			//datadir is a subdir of jobdir which contains most of the files
			DirectoryInfo datadir = jobdir.CreateSubdirectory("Data");

			//copy WMV to its local target
			//PRI3: if the user clicks Stop during this copy we have no recourse but to abort the thread, right?
			File.Copy(wmvFile,Path.Combine(datadir.FullName,job.BaseName + ".wmv"));

			//copy presentation data and slides to the local target.
			String presUrl = null;
			if (pStreamWriter != null)
			{
				presUrl = job.BaseName + ".xml";
				String presFile = pStreamWriter.WriteHeaderAndCopy(job.Segment[0].StartTime,@"./","jpg",job.ArchiveName);
				if (presFile != null)
				{
					File.Move(presFile,Path.Combine(datadir.FullName,job.BaseName + ".xml"));
				}
				if (slideImageGenerator != null)
				{
					foreach (Guid g in slideImageGenerator.OutputDirs.Keys)
					{
						DirectoryInfo diOutDir = new DirectoryInfo(Path.Combine(datadir.FullName,g.ToString()));
						diOutDir.Create();
						DirectoryInfo diSrcDir = new DirectoryInfo((String)slideImageGenerator.OutputDirs[g]);
						FileInfo[] files = diSrcDir.GetFiles();
						foreach (FileInfo f in files)
						{
							f.CopyTo(Path.Combine(diOutDir.FullName,f.Name));
						}
					}
				}
                if (this.frameGrabDirs != null) {
                    foreach (Guid g in this.frameGrabDirs.Keys) {
                        DirectoryInfo diOutDir = new DirectoryInfo(Path.Combine(datadir.FullName, g.ToString()));
                        diOutDir.Create();
                        DirectoryInfo diSrcDir = new DirectoryInfo(this.frameGrabDirs[g]);
                        FileInfo[] files = diSrcDir.GetFiles();
                        foreach (FileInfo f in files) {
                            f.CopyTo(Path.Combine(diOutDir.FullName, f.Name));
                        }
                    }
                }
			}

			//Create the asx and wbv
			writeAsx(Path.Combine(datadir.FullName,job.BaseName + ".asx"),job.BaseName + ".wmv",presUrl,job.ArchiveName); 
			writeWbv(Path.Combine(jobdir.FullName,"OpenWithWebViewer.wbv"), @"Data\" + job.BaseName + ".asx");
			
			//write readme file
			copyReadme(Path.Combine(jobdir.FullName,"ReadMe.txt"),job.ArchiveName);

			deployNotes.Append("Deployment Notes for Download Target \r\n \r\n");
			deployNotes.Append("The directory " + jobdir.Name + " contains all the files necessary " +
				"to play back the archive with WebViewer or Windows Media Player. " +
				"To deploy, simply make the directory available to your clients, for example, " +
				"zip the directory and post it on a website, or place the directory on removable media. \r\n");

			StreamWriter deploySw = File.CreateText(Path.Combine(outDir,"Deployment Notes.txt"));
			deploySw.Write(deployNotes.ToString());
			deploySw.Close();

			return null;
		}

		/// <summary>
		/// If the last element in url looks like a file name, return it without extension.
		/// Otherwise just return the default name.
		/// </summary>
		/// <param name="url"></param>
		/// <param name="defaultName"></param>
		/// <returns></returns>
		private String getBaseFileName(String url, String defaultName)
		{
			if (url == null)
				return defaultName;

			Uri u = new Uri(url.Trim());
			if (u.Segments.Length > 0)
			{
				String lastName = u.Segments[u.Segments.Length -1];
				if (lastName.IndexOf(".") > 0)
				{
					return lastName.Substring(0,lastName.LastIndexOf("."));
				}
			}
			return defaultName;
		}

		private void writeAsx(String outFileName, String wmvUrl, String presUrl, String archiveName)
		{
			StreamWriter asxfile = File.CreateText(outFileName);
			asxfile.WriteLine("<ASX version = \"3.0\">");
			asxfile.WriteLine("  <Entry>");
			asxfile.WriteLine("    <Ref href = \"" + wmvUrl + "\"/>");
			asxfile.WriteLine("    <Title> " + archiveName + " </Title>");
			//PRI2: add copyright, author ...
			if (presUrl != null)
				asxfile.WriteLine("    <PARAM name=\"CXP_SCRIPT\"  value=\"" + presUrl + "\"/>");
			asxfile.WriteLine("  </Entry>");
			asxfile.WriteLine("</ASX>"); 
			asxfile.Close();
		}

		private void writeWbv(String outFileName, string asxUrl)
		{
			StreamWriter wbvfile = File.CreateText(outFileName);
			wbvfile.WriteLine("<wmref href=\"" + asxUrl + "\"> </wmref>");
			wbvfile.Close();
		}

		private String combineUrl(String basePart, String endPart)
		{
			if (basePart.Trim().EndsWith("/"))
			{
				return basePart + endPart;
			}
			else
			{
				return basePart + "/" + endPart;
			}
		}

		/// <summary>
		/// Compose a readme.txt from an optional template and the media title.
		/// </summary>
		/// <param name="dir"></param>
		/// <param name="title"></param>
		private void copyReadme(String outPath, String title)
		{
			//Put the title at the top of the readme.
			StreamWriter readme = File.CreateText(outPath);
			readme.WriteLine("Media title: " + title);
			readme.Close();

			Assembly a = Assembly.GetExecutingAssembly();
			String exepath = Path.GetDirectoryName(a.Location);

			//Look for a template from which to copy the rest of the file.
			if (File.Exists(exepath + @"\readme_template.txt"))
			{
				FileStream readme2 =File.Open(outPath,FileMode.Append,FileAccess.Write);
				FileStream template = File.OpenRead(exepath + @"\readme_template.txt");
				byte[] buf = new byte[template.Length];
				template.Read(buf,0,(int)template.Length);
				readme2.Write(buf,0,(int)template.Length);
				template.Close();
				readme2.Close();
			}
			else
			{
				//Debug.WriteLine("Template for readme.txt does not exist at application root: " + exepath + @"\readme_template.txt");
				StreamWriter readme3 = File.AppendText(outPath);
				readme3.WriteLine(@"This content is packaged for use with WebViewer.");
				readme3.WriteLine(@"See http://www.cs.washington.edu/education/dl/confxp/webviewer.html for more information.");
				readme3.Close();
			}
			
		}

		#endregion Output Packaging

		#region Misc

        /// <summary>
        /// Return profile data from the first segment in the job that does not have the 'Slides replace video' 
        /// flag set.  If there are no such segments in the job, then construct a ProfileData created from
        /// the first segment's audio and a built-in video codec designation.
        /// </summary>
        /// <param name="job"></param>
        /// <returns></returns>
        private ProfileData jobToProfileData(ArchiveTranscoderJob job)
        {
            ProfileData pd = null;
            for (int i = 0; i< job.Segment.Length; i++)
            {
                if (Utility.SegmentFlagIsSet(job.Segment[i],SegmentFlags.SlidesReplaceVideo))
                    continue;
                else
                {
                    pd = ProfileUtility.SegmentToProfileData(job.Segment[i]);
                    break;
                }
            }
            if (pd == null)
            {
                pd = ProfileUtility.SegmentToProfileData(job.Segment[0]);
                if (pd == null) return null;
                //Specify video codec and settings to be used in lieu of an actual native profile.
                //pd.SetSecondaryVideoCodecData(WMGuids.WMCMEDIASUBTYPE_WMV3, 240, 320, 240000, 5000, 333333);
                //WMV2 works better than WMV3 for static images such as slides.
                pd.SetSecondaryVideoCodecData(WMGuids.WMCMEDIASUBTYPE_WMV2, 240, 320, 240000, 5000, 333333);
            }

            return pd;
        } 

		/// <summary>
		/// Merge RTDocuments found across all segments
		/// </summary>
		/// <param name="rtdocs"></param>
        private void accumulateRTDocuments(Hashtable rtdocs, LogMgr log)
        {
            if (rtdocs == null)
                return;

            if (rtDocuments == null)
                rtDocuments = new Hashtable();

            foreach (Guid g in rtdocs.Keys)
            {
                if (!rtDocuments.ContainsKey(g))
                    rtDocuments.Add(g, rtdocs[g]);
                else
                {
                    log.WriteLine("RTDocument with identifier " + g.ToString() + " was found in multiple segments.");
                    log.ErrorLevel = 3;
                }
            }
        }


        private void accumulateSlideMessages(Hashtable sm, LogMgr log)
        {
            if (sm == null)
                return;

            if (slideMessages == null)
                slideMessages = new Hashtable();

            foreach (String s in sm.Keys)
            {
                if (!slideMessages.ContainsKey(s))
                    slideMessages.Add(s, sm[s]);
                else
                {
                    log.WriteLine("SlideMessage with identifier " + s.ToString() + " was found in multiple segments.");
                    log.ErrorLevel = 3;
                }
            }
        }

        /// <summary>
        /// Add the current frame grab directory (if any) to the local list.
        /// </summary>
        /// <param name="frameGrabDirectory"></param>
        private void accumulateVideoStillsDirectory(WMSegment currentSegment) {
            if (currentSegment.FrameGrabDirectory != null) {
                if (this.frameGrabDirs == null) {
                    this.frameGrabDirs = new Dictionary<Guid,string>();
                }
                this.frameGrabDirs.Add(currentSegment.FrameGrabGuid,currentSegment.FrameGrabDirectory);
            }
        }

		/// <summary>
		/// ProgressTracker callback
		/// </summary>
		/// <param name="message"></param>
		private void OnShowProgress(String message)
		{
			if (this.OnStatusReport != null)
			{
				OnStatusReport(message);
			}
		}

		/// <summary>
		/// Thread proc used to stop an encoding job
		/// </summary>
		private void stopEncodingThread()
		{
			Stop();
			//Raise event to update UI.
			if (OnStopCompleted != null)
				OnStopCompleted();
		}


		/// <summary>
		/// If deck guids exist in deckTitles for this segment, but they are not found in the job, log a warning.
		/// </summary>
		/// <param name="jobLog"></param>
		/// <param name="job"></param>
		/// <param name="deckTitles"></param>
		private void logMissingDeckWarnings(LogMgr log,ArchiveTranscoderJob job,Hashtable deckTitles, int segmentNum)
		{
			if (deckTitles ==null)
				return;

			//compile a list of decks specified across all segments in the job
			ArrayList jobDecks = new ArrayList();
			if (job.Segment != null)
			{
				foreach (ArchiveTranscoderJobSegment s in job.Segment)
				{
					if (s.SlideDecks != null)
					{
						foreach (ArchiveTranscoderJobSlideDeck deck in s.SlideDecks)
						{
							if (!jobDecks.Contains(deck.DeckGuid))
							{
								jobDecks.Add(deck.DeckGuid);
							}
						}
					}
				}
			}

			//verify that all decks referenced in the segment are found in the job somewhere.
			foreach (Guid g in deckTitles.Keys)
			{
				if (!jobDecks.Contains(g.ToString()))
				{
					//Hack: this warning is generally not relevant for RTDocuments
					if ((string)deckTitles[g] != "RTDocument Deck")
					{
						log.WriteLine("Warning: Deck '" + deckTitles[g] + "' was referenced in the presentation data for segment " +
							segmentNum.ToString() + ", but not found in the job specification." +
							"  The deck id is: " + g.ToString());
						log.ErrorLevel = 3;
					}
				}
			}
		
		}

        /// <summary>
        /// If any segment has a presentation descriptor, and is not configured to use slides in place of video,
        /// return true.
        /// </summary>
        /// <param name="job"></param>
        /// <returns></returns>
		private bool containsPresentationData(ArchiveTranscoderJob job)
		{
			for(int i=0; i<job.Segment.Length; i++)
			{
				if ((job.Segment[i].PresentationDescriptor != null) && 
					(job.Segment[i].PresentationDescriptor.PresentationCname != null) && 
					(job.Segment[i].PresentationDescriptor.PresentationCname != "") &&
                    (!Utility.SegmentFlagIsSet(job.Segment[i], SegmentFlags.SlidesReplaceVideo)))
					return true;
			}
			return false;
		}

		/// <summary>
		/// Check for the existence of output directories for the specified targets.
		///		job.Path\job.BaseName\stream
		///		job.Path\job.BaseName\download
		/// </summary>
		/// <returns></returns>
		private bool outputFilesExist()
		{
			for (int i=0; i<batch.Job.Length; i++)
			{
				ArchiveTranscoderJob job = batch.Job[i];
				if (Directory.Exists(job.Path))
				{
					String streamDir = "";
					String downloadDir = "";
					if (job.Path.EndsWith(@"\"))
					{
						streamDir = job.Path + job.BaseName + @"\stream";
						downloadDir = job.Path + job.BaseName + @"\download";
					}
					else
					{
						streamDir = job.Path + @"\" + job.BaseName + @"\stream";	
						downloadDir = job.Path + @"\" + job.BaseName + @"\download";
					}

					foreach(ArchiveTranscoderJobTarget t in job.Target)
					{
						if ((t.Type.ToLower().Trim() == "stream") && (Directory.Exists(streamDir)))
						{
							return true;
						}
						if ((t.Type.ToLower().Trim() == "download") && (Directory.Exists(downloadDir)))
						{
							return true;
						}
					}
				}
			}
			return false;
		}

		/// <summary>
		/// If Constants.SQLConnectionString contains a {0}, substitute sqlHost in place.  If it contains
		/// a {1} as well, substitute dbName too.  Otherwise just use Constants.SQLConnectionString as is.
		/// Write the result to DatabaseUtility.SQLConnectionString.
		/// </summary>
		private void resetConnectionString()
		{
			if ((Constants.SQLConnectionString.IndexOf("{0}")>=0) &&
				(Constants.SQLConnectionString.IndexOf("{1}")>=0))
				DatabaseUtility.SQLConnectionString = String.Format(Constants.SQLConnectionString,sqlHost,dbName);
			else if (Constants.SQLConnectionString.IndexOf("{0}")>=0)
			{
				DatabaseUtility.SQLConnectionString = String.Format(Constants.SQLConnectionString,sqlHost);		
			} 
			else
				DatabaseUtility.SQLConnectionString = Constants.SQLConnectionString;
		}

		#endregion Misc

		#region Registry Save/Restore

		/// <summary>
		/// Get the last saved SQLHost and DBName from the registry.
		/// </summary>
		private void restoreRegSettings()
		{
			sqlHost = "";
			dbName = "ArchiveService";
			try
			{
				RegistryKey BaseKey = Registry.CurrentUser.OpenSubKey(Constants.AppRegKey, true);
				if ( BaseKey == null) 
				{ //no configuration yet.. first run.
					//Debug.WriteLine("No registry configuration found.");
					return;
				}
				sqlHost = Convert.ToString(BaseKey.GetValue("SQLHost",sqlHost));
				dbName = Convert.ToString(BaseKey.GetValue("DBName",dbName));
			}
			catch (Exception e)
			{
				Debug.WriteLine("Exception while reading registry: " + e.ToString());
			}
		}

		/// <summary>
		/// Store SqlHost and DbName in the registry
		/// </summary>
		private void saveRegSettings()
		{
			try
			{
				RegistryKey BaseKey = Registry.CurrentUser.OpenSubKey(Constants.AppRegKey, true);
				if ( BaseKey == null) 
				{
					BaseKey = Registry.CurrentUser.CreateSubKey(Constants.AppRegKey);
				}
				BaseKey.SetValue("SQLHost",sqlHost);
				BaseKey.SetValue("DBName",dbName);
			}
			catch
			{
				Debug.WriteLine("Exception while saving configuration.");
			}
		}

		#endregion Registry Save/Restore

		#endregion Private Methods

		#region Events

		/// <summary>
		/// Report progress status.
		/// </summary>
		public event statusReportHandler OnStatusReport;
		public delegate void statusReportHandler(String message);

		/// <summary>
		/// Work is completed.
		/// </summary>
		public event batchCompletedHandler OnBatchCompleted;
		public delegate void batchCompletedHandler();

		/// <summary>
		/// The AsyncStop thread has completed.
		/// </summary>
		public event stopCompletedHandler OnStopCompleted;
		public delegate void stopCompletedHandler();

		#endregion

        #region Testing

        /// <summary>
        /// One way to run this is using the -t option to the console app.
        /// </summary>
        public void RunUnitTests() {
            //Test.VerifyStream(48, DateTime.Parse("4/8/2013 8:20:00 PM"), DateTime.Parse("4/8/2013 8:47:36 PM"));
            Test.UnPillarBoxTest();
        }


        /// <summary>
        /// It's assumed that the job contains only 'presentation from video' segments
        /// Don't build any WMV, just build the presentation and experiment with filtering images.
        /// </summary>
        public void FrameGrabTest() {
            ArchiveTranscoderJob job = this.batch.Job[0];
            LogMgr log = new LogMgr();
            ProgressTracker progressTracker = new ProgressTracker(job.Segment.Length);
            progressTracker.Start();
            List<WMSegment> segs = new List<WMSegment>();
            for (int j = 0; j < job.Segment.Length; j++) {
                WMSegment s = new WMSegment(job, job.Segment[j], 0, 0, wmWriter, progressTracker, log, null, true, null);
                segs.Add(s);
                s.FrameGrabTest();
            }
            PresentationFromVideoMgr.FilterSegments(segs);
            progressTracker.Dispose();
        }

        #endregion

    }
}
