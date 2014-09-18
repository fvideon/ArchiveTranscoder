using System;
using System.Diagnostics;
using System.Collections;
using System.Threading;
using MSR.LST;
using MSR.LST.Net.Rtp;
using MSR.LST.RTDocuments;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using System.Text.RegularExpressions;
using CP3 = UW.ClassroomPresenter;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Perform Presentation data analysis to determine which slide decks are used.  Provide services to the UI
	/// to support matching referenced decks with actual decks.  An instance of the class corresponds with an
	/// instance of the DeckMatcherForm.
	/// </summary>
	public class DeckMatcher
	{
		#region Members

		private String cname;
		private String start;
		private String end;
		private DateTime dtstart;
		private DateTime dtend;
		private Hashtable decks;
        private Hashtable unnamedDecks;
		private bool autoMatchDone;
		private Thread autoMatchThread;
		private Thread analyzeThread;
		private bool stopAutoMatch;
		private bool stopAnalyze;
		private DirectoryInfo autoMatchDir;
		private ProgressTracker progressTracker;
		private bool pptInstalled;
		private PayloadType payload;

		#endregion Members

		#region Property

		/// <summary>
		/// Contains decks found, keyed by deck Guid.
		/// </summary>
		public Hashtable Decks
		{
			get { return decks; }
		}

		#endregion Property

		#region Ctor/Dispose

		public DeckMatcher(String cname, String payload, String start, String end, bool pptInstalled)
		{
			this.cname = cname;
			this.start = start;
			this.end = end;
			this.payload = Utility.StringToPayloadType(payload);
			this.pptInstalled = pptInstalled;
			decks = new Hashtable();
		}

		#endregion Ctor/Dispose

		#region Public Methods

		/// <summary>
		/// Make sure all the threads are ended.
		/// </summary>
		public void StopThreads()
		{
			if (autoMatchThread != null)
			{
				stopAutoMatch = true;
				if (!autoMatchThread.Join(1000))
				{
					autoMatchThread.Abort();
				}
				autoMatchThread = null;
			}

			if (analyzeThread != null)
			{
				stopAnalyze = true;
				if (!analyzeThread.Join(1000))
				{
					analyzeThread.Abort();
				}
				analyzeThread = null;
			}

			if (progressTracker != null)
			{
				progressTracker.Stop();
				progressTracker = null;
			}
		}

		/// <summary>
		/// Remove the specified deck from the Hashtable
		/// </summary>
		/// <param name="d"></param>
		public void Remove(Deck d)
		{
			if (decks.ContainsKey(d.DeckGuid))
			{
				decks.Remove(d.DeckGuid);
			}
		}


		/// <summary>
		/// Start the thread to find out which decks are referenced in the presentation data.
		/// </summary>
		/// <returns></returns>
		public String Analyze()
		{
			dtstart = DateTime.MinValue;
			dtend = DateTime.MinValue;
			try
			{
                //Bug: we should start looking from the beginning of the stream to make sure we find decks in all cases.
                // in particular CP3 decks won't be found if there wasn't a transition in the selected time range.
				dtstart = DateTime.Parse(start);
			}
			catch 
			{
				return("Failed to parse segment start time as a time: " + start);
			}

			try
			{
				dtend = DateTime.Parse(end);
			}
			catch 
			{
				return("Failed to parse segment end time as a time: " + end);

			}

			if (dtstart >= dtend)
			{
				return("Start time comes after or is equal to end time!");
			}

			if ((cname == null) || (cname == ""))
			{
				return("Presentation source is not specified.");
			}

			if (analyzeThread !=null)
			{
				if (analyzeThread.IsAlive)
				{
					return "already running analyze";
				}
			}

			stopAnalyze = false;
			analyzeThread = new Thread(new ThreadStart(AnalyzeThread));
			analyzeThread.Name = "DeckMatcher Analyze Thread";
			analyzeThread.Start();
			return null;
		}

		/// <summary>
		/// Stop analyzing presentation data.
		/// </summary>
		public void StopAnalyze()
		{
			if (analyzeThread == null)
			{
				return;
			}
			stopAnalyze = true;
			if (!analyzeThread.Join(2000))
			{
				Debug.WriteLine("Deckmatcher analysis thread aborting.");
				analyzeThread.Abort();
			}
			if (progressTracker != null)
			{
				progressTracker.Stop();
				progressTracker = null;
			}
			analyzeThread = null;
		}

		/// <summary>
		/// Use to prepopulate the DeckMatcher with persisted deck info.
		/// </summary>
		/// <param name="deck"></param>
		public void Add(Deck deck)
		{
			if (!decks.ContainsKey(deck.DeckGuid))
			{
				decks.Add(deck.DeckGuid,deck);
			}
		}

		/// <summary>
		/// Begin a thread to search for appropriate decks in the filesystem tree rooted at the given directory
		/// </summary>
		/// <param name="rootDir"></param>
		/// <returns></returns>
		public String AutoMatch(DirectoryInfo rootDir)
		{
			if (autoMatchThread !=null)
			{
				if (autoMatchThread.IsAlive)
				{
					return "already running auto-match";
				}
			}

			autoMatchDir = rootDir;
			stopAutoMatch = false;
			autoMatchThread = new Thread(new ThreadStart(AutoMatchThread));
			autoMatchThread.Name = "DeckMatcher Auto-Match Thread";
			autoMatchThread.Start();
			return null;
		}

		/// <summary>
		/// Stop the auto-match thread
		/// </summary>
		public void StopAutoMatch()
		{
			if (autoMatchThread == null)
			{
				return;
			}
			stopAutoMatch = true;
			if (!autoMatchThread.Join(2000))
			{
				autoMatchThread.Abort();
			}
			autoMatchThread = null;
		}

		#endregion Public Methods

		#region Private Methods 

		/// <summary>
		/// Thread proc to examine presentation data to find referenced slide decks.
		/// </summary>
		private void AnalyzeThread()
		{
			decks.Clear();
            this.unnamedDecks = new Hashtable();
			BufferChunk frame;
			long timestamp;

			progressTracker = new ProgressTracker(1);
			progressTracker.CustomMessage = "Initializing";
			progressTracker.CurrentValue=0;
			progressTracker.EndValue=100; //dummy value to start with.
			progressTracker.OnShowProgress += new ProgressTracker.showProgressCallback(OnShowProgress);
			progressTracker.Start();

			//Note: there would be multiple streams here if the presenter left the venue and rejoined during the segment.
			// This is in the context of a single segment, so we can assume one cname, and no overlapping streams.
			int[] streams = DatabaseUtility.GetStreams(payload,cname,null,dtstart.Ticks,dtend.Ticks);
			if (streams==null)
			{
				progressTracker.Stop();
				progressTracker = null;
				this.OnAnalyzeCompleted();
				return;
			}

			DBStreamPlayer[] streamPlayers = new DBStreamPlayer[streams.Length];
			double totalDuration = 0;
			for(int i=0; i<streams.Length; i++)
			{
				streamPlayers[i] = new DBStreamPlayer(streams[i],dtstart.Ticks,dtend.Ticks,this.payload);
				totalDuration += ((TimeSpan)(streamPlayers[i].End - streamPlayers[i].Start)).TotalSeconds;
			}
			
			progressTracker.EndValue = (int)totalDuration;
			progressTracker.CustomMessage = "Analyzing";

			int framecount = 0;

			for (int i=0;i<streamPlayers.Length;i++)
			{
				if (stopAnalyze)
					break;
				while ((streamPlayers[i].GetNextFrame(out frame,out timestamp)))
				{
					if (stopAnalyze)
						break;
					ProcessFrame(frame);
					progressTracker.CurrentValue = (int)(((TimeSpan)(new DateTime(timestamp) - streamPlayers[0].Start)).TotalSeconds);
					framecount++;
				}
			}

			if ((!stopAnalyze) && (this.OnAnalyzeCompleted != null))
			{
                mergeUnnamedDecks();
				progressTracker.Stop();
				progressTracker = null;
				this.OnAnalyzeCompleted();
			}
		}

        private void mergeUnnamedDecks() {
            foreach (Guid g in this.unnamedDecks.Keys) {
                if (!decks.ContainsKey(g)) {
                    decks.Add(g, unnamedDecks[g]);
                    if (this.OnDeckFound != null) {
                        OnDeckFound((Deck)this.unnamedDecks[g]);
                    }
                }
            }
        }



		/// <summary>
		/// Thread proc to search for appropriate decks to match referenced decks.
		/// For now just look for CSD files under the root dir whose names match unmatched decks.
		/// return true if one or more decks were matched.
		/// </summary>
		/// <param name="rootDir"></param>
		private void AutoMatchThread()
		{
			//clear existing matches.
			foreach (Deck d in decks.Values)
			{
				d.Path = null;
				d.Matched = false;
			}

			autoMatchDone = false;
			autoMatch(autoMatchDir);
			
			if ((!stopAutoMatch) && (this.OnAutoMatchCompleted != null))
			{
				OnAutoMatchCompleted();
			}
		}

		private int countMatches()
		{
			int matchCount = 0;
			foreach(Deck d in decks.Values)
			{
				if (d.Matched)
					matchCount ++;
			}
			return matchCount;
		}

		/// <summary>
		/// Recursively search a directory for supported deck files whose names match the decks we know about.
        /// A complete match overrides a partial match.  PPT or PPTX override any.  CP3 overrides CSD.
		/// </summary>
		/// <param name="rootDir"></param>
		private void autoMatch(DirectoryInfo rootDir)
		{
			if ((autoMatchDone) || (stopAutoMatch))
				return;

            FileInfo[] files = null;
            try {
                files = rootDir.GetFiles();
            }
            catch {
                return; //Most likely a permission issue.
            }

			foreach(FileInfo fi in files) {  //iterate over files

                string thisExt = fi.Extension.ToLower();

                if (!((thisExt.Equals(".ppt")) ||
                    (thisExt.Equals(".pptx")) ||
                    (thisExt.Equals(".cp3")) ||
                    (thisExt.Equals(".csd"))
                    )) {
                    //This file is not a supported deck type.
                    continue;
                }

                if (((thisExt.Equals(".ppt")) ||
                    (thisExt.Equals(".pptx"))) && (!pptInstalled)) { 
                    //PPT not supported on this system; skip this one.
                    continue;
                }

                bool allMatched = true;

				foreach (Deck d in decks.Values) { //iterate over decks comparing each to the current file

                    if ((d.ExactMatch) && 
                        ((d.MatchExt.Equals(".ppt")) || (d.MatchExt.Equals(".pptx")) || (!pptInstalled))) {
                        //Already have an optimal match for this one
                        continue;
                    }

                    allMatched = false; //No existing optimal match for this Deck d.

                    bool exact = false;
                	String deckBase = StripNumInBrackets(d.FileName);
                    if (Path.GetFileNameWithoutExtension(fi.Name).Equals(deckBase, StringComparison.CurrentCultureIgnoreCase)) {
                        exact = true; //exact match
                    }
                    else if (fi.Name.IndexOf(deckBase, StringComparison.CurrentCultureIgnoreCase) >= 0) {
                        //partial match
                    }
                    else {
                        //no match
                        continue;
                    }

                    //here we have at least a partial match between the current file and the current deck.
                    //Record the match if it is better than the match already recorded (if any).

                    if (betterMatch(exact, thisExt, d.ExactMatch, d.MatchExt)) { 
                        d.MatchExt = thisExt;
                        d.Matched = true;
                        d.Path = fi.FullName;
                        d.ExactMatch = exact;                   
                    }

				}

                if (allMatched) {
                    //We have optimal matches for all decks.
                    autoMatchDone = true;
                    return;
                }

			}

            //Recurse for each subdirectory.
			foreach (DirectoryInfo di in rootDir.GetDirectories())
			{
				autoMatch(di);
                if (autoMatchDone || stopAutoMatch) break;
			}
		}

        /// <summary>
        /// Figure out if the new match is better than the existing match (if any)
        /// </summary>
        /// <param name="newExact">New match is exact</param>
        /// <param name="newExt">Extension of new file</param>
        /// <param name="oldExact">Existing match is exact</param>
        /// <param name="oldExt">Extension of existing match file</param>
        /// <returns>True if new match is better</returns>
        private bool betterMatch(bool newExact, string newExt, bool oldExact, string oldExt) {
            if (oldExt == null) { //No existing match.  Any match overrides this.
                return true;
            }
            else if (!oldExact) { //existing match is partial
                if (newExact) { //new exact match overrides any partial match
                    return true;
                }
                else { //both existing and new matches are partial.
                    //return true if new has a perferred file type
                    return betterExtension(newExt, oldExt);
                }
            }
            else if (newExact) { //There is already an exact match and current match is exact
                //return true if new has a preferred file type
                return betterExtension(newExt, oldExt);
            }

            return false;
        }

        /// <summary>
        /// Determine if the new file type is preferred when forming a match.
        /// </summary>
        /// <param name="newExt">new File extension</param>
        /// <param name="oldExt">existing file extension</param>
        /// <returns>true if new file type is preferred.</returns>
        private bool betterExtension(string newExt, string oldExt) {
            if (newExt.Equals(".csd")) {
                //CSD overrides nothing
            }
            else if (newExt.Equals(".cp3")) {
                if (oldExt.Equals(".csd")) { //CP3 overrides CSD
                    return true;
                }
            }
            else if ((newExt.Equals(".ppt")) || (newExt.Equals(".pptx"))) {
                //ppt(x) overrides any but ppt(x).
                if (!((oldExt.Equals(".ppt")) || (oldExt.Equals(".pptx")))) {
                    return true;
                }
            }
            return false;
        }

		/// <summary>
		/// if there is an integer in [] at the end of the string, strip it off.
		/// </summary>
		/// <param name="s"></param>
		/// <returns></returns>
		private static String StripNumInBrackets(String s)
		{
			Regex re = new Regex(@"^(?<1>.*)\[\d+\]$");
			Match m = re.Match(s);
			if (m.Success)
			{
				return m.Groups[1].ToString();
			}
			return s;
		}

		/// <summary>
		/// Examine a frame in the stream, looking for deck references.
		/// </summary>
		/// <param name="frame"></param>
		private void ProcessFrame(BufferChunk frame)
		{
			BinaryFormatter bf = new BinaryFormatter();
			Object rtobj = null;
			try 
			{
				MemoryStream ms = new MemoryStream((byte[]) frame);
				rtobj = bf.Deserialize(ms);
			}
			catch (Exception e)
			{
				Debug.WriteLine("ProcessFrame exception deserializing message. size=" + frame.Length.ToString() +
					" exception=" + e.ToString());
				return;
			}

            if (payload == PayloadType.dynamicPresentation) {
                if (rtobj is PresenterNav.CPDeckCollection) {
                    addDecks((PresenterNav.CPDeckCollection)rtobj);
                }
                else if (rtobj is CP3.Network.Chunking.Chunk) { 
                    //CP3 
                    addCp3Decks((CP3.Network.Chunking.Chunk)rtobj);
                }
            }
            else if (payload == PayloadType.RTDocument) {
                if (rtobj is RTFrame) {
                    RTFrame rtFrame = (RTFrame)rtobj;
                    if (rtFrame.ObjectTypeIdentifier == Constants.CLASSROOM_PRESENTER_MESSAGE) //This is CP used as a capability
					{
                        if (rtFrame.Object is PresenterNav.CPDeckCollection) {
                            addDecks((PresenterNav.CPDeckCollection)rtFrame.Object);
                        }
                    }
                }
                else {
                    //As far as I know, the RTDocument message is the only sure way to tell that there is a slide deck.
                    // other messages contain a TOC identifer that maps to the RTDocument, so by themselves they don't provide
                    // enough info.  We will need to be able to find the RTDocument in the data when it comes time to 
                    // process, and we will assume that slide images will also be in the data, so I think the best thing
                    // to do is to forget about deck matching in this case.
                }
            }

		}


        private void addCp3Decks(UW.ClassroomPresenter.Network.Chunking.Chunk chunk) {
            if (chunk.NumberOfChunksInMessage == 1) {
                MemoryStream ms = new MemoryStream(chunk.Data);
                BinaryFormatter bf = new BinaryFormatter();
                Object o = bf.Deserialize(ms);
                if (o is CP3.Network.Messages.Message) {
                    ProcessCP3MessageGraph((CP3.Network.Messages.Message)o);
                }
            }
            else {
                Debug.WriteLine("Warning: DeckMatcher does not yet support multi-chunk messages");
            }
        }

        private void ProcessCP3MessageGraph(UW.ClassroomPresenter.Network.Messages.Message m) {
            if (m.Predecessor != null) {
                ProcessCP3MessageGraph(m.Predecessor);
            }

            ProcessCP3Message(m);

            if (m.Child != null) {
                ProcessCP3MessageGraph(m.Child);
            }

        }

        private void ProcessCP3Message(CP3.Network.Messages.Message m) {
            if (m is CP3.Network.Messages.Presentation.DeckInformationMessage) {
                CP3.Network.Messages.Presentation.DeckInformationMessage dim = (CP3.Network.Messages.Presentation.DeckInformationMessage)m;
                if ((dim.Disposition == UW.ClassroomPresenter.Model.Presentation.DeckDisposition.Whiteboard) ||
                    (dim.Disposition == UW.ClassroomPresenter.Model.Presentation.DeckDisposition.StudentSubmission) ||
                    (dim.Disposition == UW.ClassroomPresenter.Model.Presentation.DeckDisposition.QuickPoll)) {
                    return;
                }
                if (!decks.ContainsKey(dim.TargetId)) {
                    Deck thisDeck = new Deck(dim);
                    decks.Add(dim.TargetId, thisDeck);
                    if (this.OnDeckFound != null) {
                        OnDeckFound(thisDeck);
                    }
                }
            }
            else if (m is CP3.Network.Messages.Network.InstructorCurrentDeckTraversalChangedMessage) {
                CP3.Network.Messages.Network.InstructorCurrentDeckTraversalChangedMessage im = (CP3.Network.Messages.Network.InstructorCurrentDeckTraversalChangedMessage)m;
                if (((im.Dispositon & UW.ClassroomPresenter.Model.Presentation.DeckDisposition.Whiteboard) == 0) && 
                    (!this.unnamedDecks.ContainsKey(im.DeckId))) {
                    // This case captures all the decks, including cases where the deck is opened before
                    // the start of archiving, and there are no slide transitions within the deck, but 
                    // one slide is shown.  These need to be manually matched since we don't know the deck name.
                    // At the end of processing we merge these into the main decks collection.
                    this.unnamedDecks.Add(im.DeckId, new Deck(im.DeckId));
                }
            }
        }


		private void addDecks(PresenterNav.CPDeckCollection cpdc)
		{
			foreach (PresenterNav.CPDeckSummary cpds in cpdc.SummaryCollection)
			{
				if ((cpds.DeckType=="Presentation") && (!decks.ContainsKey(cpds.DocRef.GUID)))
				{
					Deck thisDeck = new Deck(cpds);
					decks.Add(cpds.DocRef.GUID,thisDeck);
					if (this.OnDeckFound != null)
					{
						OnDeckFound(thisDeck);
					}
				}
			}
		}


		/// <summary>
		/// ProgressTracker event handler.
		/// </summary>
		/// <param name="message"></param>
		private void OnShowProgress(String message)
		{
			if (this.OnStatusReport != null)
			{
				OnStatusReport(message);
			}
		}

		#endregion Private Methods

		#region Debugging & Data Analysis

		private void DumpPageUpdate(PresenterNav.CPPageUpdate cppu)
		{
			Debug.WriteLine("PageUpdate packet dump: \n" +
				"  ViewportIndex=" + cppu.ViewPortIndex.ToString() + "\n" +
				"  DeckType=" + cppu.DeckType + "\n" +
				"  FileName=" + cppu.FileName + "\n" +
				"  PageRef.GlobalID=" + cppu.PageRef.GlobalID.ToString() + "\n" +
				"  PageRef.LocalID=" + cppu.PageRef.LocalID.ToString() + "\n" +
				"  PageRef.Index=" + cppu.PageRef.Index.ToString() +  "\n" +
				"  DocRef.DocName=" + cppu.DocRef.DocName + "\n" +
				"  DocRef.GUID=" + cppu.DocRef.GUID.ToString() + "\n" +
				"  DocRef.Index=" + cppu.DocRef.Index.ToString());
		}

		private void DumpScrollLayer(ArchiveRTNav.RTScrollLayer rtsl)
		{
			Debug.WriteLine("ScrollLayer: si=" + rtsl.SlideIndex.ToString() +
				" se=" + rtsl.ScrollExtent.ToString() +
				" sp=" + rtsl.ScrollPosition.ToString() +
				" di=" + rtsl.DeckGuid.ToString()); 
		}
		private void DumpScrollLayer(PresenterNav.CPScrollLayer cpsl)
		{
			Debug.WriteLine("ScrollLayer packet dump: \n" +
				"  DeckType=" + cpsl.DeckType + "\n" + 
				"  DocRef.Index=" + cpsl.DocRef.Index.ToString() + "\n" + 
				"  DocRef.DocName=" + cpsl.DocRef.DocName + "\n" + 
				"  DocRef.GUID=" + cpsl.DocRef.GUID.ToString() + "\n" + 
				"  LayerIndex=" + cpsl.LayerIndex.ToString() + "\n" + 
				"  PageRef.GlobalID=" + cpsl.PageRef.GlobalID.ToString() + "\n" + 
				"  PageRef.Index=" + cpsl.PageRef.Index.ToString() + "\n" + 
				"  PageRef.LocalID=" + cpsl.PageRef.LocalID.ToString() + "\n" + 
				"  ScrollExtent=" + cpsl.ScrollExtent.ToString() + "\n" + 
				"  ScrollPosition=" + cpsl.ScrollPosition.ToString() );
		}

		private void DumpBeacon(WorkSpace.BeaconPacket bp)
		{
			Debug.WriteLine("BeaconPacket dump: \n" +
				"  Alive=" + bp.Alive.ToString() + "\n" +
				"  BGColor=" + bp.BGColor.ToString() + "\n" +
				"  ID=" + bp.ID.ToString() + "\n" +
				"  Role=" + bp.Role.ToString() + "\n" +
				"  Version=" + bp.Version + "\n" +
				"  FName=" + bp.FriendlyName + "\n" +
				"  Name=" + bp.Name + "\n" +
				"  Time=" + bp.Time.ToString() + "\n\n"  );
		}

		private void DumpDeckCollection(PresenterNav.CPDeckCollection cpdc)
		{
			Debug.WriteLine("DeckCollection packet dump:");
			Debug.WriteLine("  StudentNavigation=" + cpdc.StudentNavigation );
			Debug.WriteLine("  StudentSubmission=" + cpdc.StudentSubmission.ToString() );
			foreach (PresenterNav.CPDeckSummary cs in cpdc.SummaryCollection)
			{
				Debug.WriteLine("  DeckSummary: \n" + 
					"    slidecount=" + cs.SlideCount.ToString() + "\n" + 
					"    decktype=" + cs.DeckType + "\n" + 
					"    deckindex=" + cs.DocRef.Index.ToString() + "\n" + 
					"    deckguid=" + cs.DocRef.GUID.ToString() + "\n" + 
					"    slidetitles.Count=" + cs.SlideTitles.Count.ToString() );
			}
			Debug.WriteLine("  DeckIndices:");
			foreach (int i in cpdc.DeckIndices)
			{	
				Debug.WriteLine("    index=" + i.ToString());
			}
			Debug.WriteLine("  RemovedDecks:");
			foreach (int i in cpdc.RemovedDecks)
			{	
				Debug.WriteLine("    index=" + i.ToString());
			}
			Debug.WriteLine("  ViewableSlides:");
			if (cpdc.ViewableSlides == null)
			{
				Debug.WriteLine("    is null");
			}
			else
			{
				foreach (int i in cpdc.ViewableSlides)
				{	
					Debug.WriteLine("    index=" + i.ToString());
				}
			}
			Debug.WriteLine("  ViewPort: \n" +
				"    SlideSize=" + cpdc.ViewPort.SlideSize.ToString() + "\n" +
				"    X=" + cpdc.ViewPort.X.ToString() + "\n" +
				"    Y=" + cpdc.ViewPort.Y.ToString());
		}

		#endregion Debugging & Data Analysis

		#region Events

		/// <summary>
		/// Analysis thread completed.  The event will not be raised if the
		/// thread was terminated by the user.
		/// </summary>
		public event analyzeCompletedHandler OnAnalyzeCompleted;
		public delegate void analyzeCompletedHandler();
		
		/// <summary>
		/// Auto-match thread completed.  The event will not be raised if the
		/// thread was terminated by the user.
		/// </summary>
		public event autoMatchCompletedHandler OnAutoMatchCompleted;
		public delegate void autoMatchCompletedHandler();

		/// <summary>
		/// Update to analysis thread 'percent complete' progress
		/// </summary>
		public event statusReportHandler OnStatusReport;
		public delegate void statusReportHandler(String message);

		/// <summary>
		/// The analysis thread found a new deck reference.
		/// </summary>
		public event deckFoundHandler OnDeckFound;
		public delegate void deckFoundHandler(Deck deck);

		#endregion Events
	}

	#region Deck Class

	public class Deck
	{
		private Guid deckGuid;
		private int slideCount;
		private String fileName;
		private String[] slideTitles;
		private bool matched;
		private String path;
        private String matchExt;
        private bool exactMatch;

		public Deck(PresenterNav.CPDeckSummary cpds)
		{
			matched = false;
            matchExt = null;
            exactMatch = false;
			deckGuid = cpds.DocRef.GUID;
			slideCount = cpds.SlideCount;
			fileName = cpds.FileName;
			slideTitles = (String[])cpds.SlideTitles.ToArray(typeof(String));
		}

        public Deck(CP3.Network.Messages.Presentation.DeckInformationMessage dim) {
            matched = false;
            matchExt = null;
            exactMatch = false;
            deckGuid = (Guid)dim.TargetId;
            slideCount = 0;
            fileName = dim.HumanName;
            slideTitles = null;
        }

        public Deck(Guid g) {
            matched = false;
            matchExt = null;
            exactMatch = false;
            deckGuid = g;
            slideCount = 0;
            fileName = "Unknown Deck ID:" + g.ToString();
            slideTitles = null;
        }

		public Deck (RTDocument rtd)
		{
			matched = false;
            matchExt = null;
            exactMatch = false;
			deckGuid = rtd.Identifier;
			slideCount = rtd.Organization.TableOfContents.Count;
			fileName = "unknown name";
			slideTitles = null;

		}

		public Deck (ArchiveTranscoderJobSlideDeck deck)
		{
			matched = false;
            matchExt = null;
            exactMatch = false;
            deckGuid = new Guid(deck.DeckGuid);
			slideCount = -1;
			fileName = deck.Title;
			if (deck.Path != null)
			{
				path = deck.Path;
				matched = true;
			}
			slideTitles = null;
		}

		public override string ToString()
		{
			String ret;
			if (matched)
				ret = "Matched Deck: ";
			else
				ret = "Unmatched Deck: ";

			ret += fileName + " ";
			if (slideCount > 0)
			{
				ret += "(" + slideCount.ToString() + " slides)";
			}

			if (matched)
				ret += " Path: " + path;

			return ret;
		}

		public ArchiveTranscoderJobSlideDeck ToSlideDeck()
		{
			ArchiveTranscoderJobSlideDeck deck = new ArchiveTranscoderJobSlideDeck();
			deck.DeckGuid = deckGuid.ToString();
			deck.Title = fileName;
			if (matched)
				deck.Path = path;
			return deck;
		}

		public bool Matched
		{
			get { return matched; }
			set { matched = value; }
		}

		public String Path
		{
			get { return path; }
			set { path = value; }
		}

        public bool ExactMatch {
            get { return exactMatch; }
            set { exactMatch = value; }
        }

        public String MatchExt {
            get { return matchExt; }
            set { matchExt = value; }
        }

		public Guid DeckGuid
		{
			get { return deckGuid; }
		}

		public int SlideCount
		{
			get { return slideCount; }
		}
				
		public String FileName
		{
			get { return fileName; }
		}
				
		public String[] SlideTitles
		{
			get { return slideTitles; }
		}
	}

	#endregion Deck Class
}
