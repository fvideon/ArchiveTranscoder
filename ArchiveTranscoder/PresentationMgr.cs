using System;
using System.IO;
using System.Collections;
using System.Diagnostics;
using System.Text;
using System.Runtime.Serialization.Formatters.Binary;
using MSR.LST.Net.Rtp;
using MSR.LST;
using MSR.LST.RTDocuments;
using ArchiveRTNav;
using WorkSpace;
using UW.ClassroomPresenter;
using System.Collections.Generic;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Given a cname, a supported presentation data type (eg. CPNav, RTDocument, etc.), and a timespan, 
	/// filter and convert relevant frames.  Produce output appropriate for writing to XML file.
	/// </summary>
	public class PresentationMgr
	{
		#region Members

		private String cname;		//Identifies the participant which sourced the presentation data for this instance
		private long start;			//absolute start time for this instance
		private long end;			//absolute end time for this instance
        private long m_ConferenceStart; //Start time for the conference to support look-behind.
		private bool stopnow;		//Flag to indicate that processing should terminate now.
		private long offset;		//Ticks to be subtracted from absolute timestamp.
		private RTUpdate rtUpdate;	//"beacon" object for output.
		private int lastSlideIndex=-1;					//Most recently reported slide index.
		private Guid lastDeckGuid=Guid.Empty;			//Most recently reported slide deck identifier
		private ScrollPosCache scrollPosCache;			//Remember scroll positions for different slides
		private ArrayList dataList;						//The output 
		private StudentSubmissionCache studentSubmissionCache;  //Tracks student submission data
        private SlideMessageCache slideMessageCache;    //Tracks certain CP slide broadcasts.
		private Hashtable seenStudentSubmissions;		//A list of student submissions for which we have already output ink
		private double lastScrollExtent = 0;			//Most recently reported scroll bounds
		private double lastScrollPosition = 0;			//Most recently reported scroll position
		private SlideTitleMonitor slideTitleMonitor;	//Gather advertised slide titles.
		private PresenterWireFormatType format;			//Wire format such as 'RTDocuments' or 'CPNav' for current instance.
		private PayloadType payload;					//RTP payload inferred from the format.
		private Hashtable rtDocs;						//Collection of RTDocuments seen.  Key is RTDocument Guid, value is RTDocument.
		private Hashtable unresolvedPages;				//RTDocument pages that don't initially match a RTDocument.
		private RTDocument wbRtDoc;						//an RTDocument to hold the Whiteboard pages
		private RTDocument miscRtDoc;					//An RTDocument to hold misc pages such as screenshot pages.
		private bool firstRTDocFrame;					//Flag to show that we have yet to emit an initial update message in the RTDoc scenairo.
        private CP3Manager.CP3Manager cp3Mgr;           //Transcode CP3 messages		
        private LogMgr errorLog;
        private List<object> initialState;              //Accumulate messages needed to establish initial presentation state
        private long m_PreviousSegmentEnd;
        private ArchiveTranscoderJobSegmentVideoDescriptor videoDescriptor; //For building a presentation from video stills.
        private PresentationFromVideoMgr presFromVideoMgr;                  //For building a presentation from video stills.
        #endregion Members

		#region Properties

		/// <summary>
		/// True indicates that this instance has found some qualified presentation data, and has a non-null output
		/// </summary>
		public bool ContainsData
		{
			get
			{
                if (dataList.Count > 0) {
                    return true;
                }
                if ((this.presFromVideoMgr != null) && 
                    (this.presFromVideoMgr.TempDir != null)){
                    return true;
                }

                return false;
			}
		}

		/// <summary>
		/// A collection of Deck Titles discovered in the presentation data, or null if none.
		/// </summary>
		public Hashtable DeckTitles
		{
			get 
			{
				if (slideTitleMonitor != null)
					return slideTitleMonitor.DeckTitles;
				else
					return null;
			}
		}

		/// <summary>
		/// A collection of RTDocuments found in the presentation streams, keyed by RTDocument Identifier (Guid)
		/// </summary>
		public Hashtable RTDocuments 
		{
			get {return rtDocs;}
		}

        /// <summary>
        /// This could contain all the broadcasts we hear, but for now it just has the QuickPoll results slides.
        /// </summary>
        public Hashtable SlideMessages
        { 
            get {return slideMessageCache.Table;}
        }

        public PresentationFromVideoMgr FrameGrabMgr {
            get { return this.presFromVideoMgr;  }
        }

		#endregion Properties

		#region Ctor

		/// <summary>
		/// Converts raw presentation streams to archive streams.
		/// </summary>
		/// <param name="presDescriptor">User specified presentation stream set</param>
		/// <param name="start">actual segment start time</param>
		/// <param name="end">actual segment end time</param>
		/// <param name="offset">offset is the cumulative gap time between segments.. 0 for the first segment. This number is 
		/// subtracted from the database timestamp to make an adjusted absolute time.</param>
		public PresentationMgr(ArchiveTranscoderJobSegmentPresentationDescriptor presDescriptor, 
			long start, long end, long offset, LogMgr errorLog, long previousSegmentEnd, ProgressTracker progressTracker)
		{
            this.errorLog = errorLog;
			stopnow = false;
			this.cname = presDescriptor.PresentationCname;
			this.format = Utility.StringToPresenterWireFormatType(presDescriptor.PresentationFormat);
            this.videoDescriptor = presDescriptor.VideoDescriptor;
            if ((this.videoDescriptor != null) && (this.format == PresenterWireFormatType.Video)) {
                this.presFromVideoMgr = new PresentationFromVideoMgr(this.videoDescriptor, start, end, offset, errorLog, previousSegmentEnd, progressTracker);
            }
			payload = Utility.formatToPayload(format);
			this.start = start;
			this.end = end;
			this.offset = offset;
            this.m_PreviousSegmentEnd = previousSegmentEnd;
			rtUpdate = new RTUpdate();
			rtUpdate.SlideSize = 1.0; //Note, not setting this can cause webviewer to crash if the archive begins at the wrong spot.
			//PRI2: Rather than setting an initial value that may be wrong, we should probably look behind to find out what it should be.
			scrollPosCache = new ScrollPosCache();
			studentSubmissionCache = new StudentSubmissionCache();
            slideMessageCache = new SlideMessageCache();
			dataList = new ArrayList();
			seenStudentSubmissions = new Hashtable();
			slideTitleMonitor = new SlideTitleMonitor();

            //Get the start time for the entire conference to support look-behind.
            m_ConferenceStart = DatabaseUtility.GetConferenceStartTime(this.payload, this.cname, this.start, this.end);
            Debug.Assert(m_ConferenceStart>0,"No conference exists in the database that matches this presentation: " + cname +
                    " with PresentationFormat " + format.ToString());
                
		}

		#endregion Ctor

		#region Public Methods

		/// <summary>
		/// Read data and process it.  Return error message if any, or null if successful.
		/// Note that the caller is responsible for writing header and footer info.
		/// This must run before Print.
		/// </summary>
		/// <returns></returns>
		/// Terminate asap if stopnow becomes true.
		public String Process()
		{
            string result = null;
			if ((format==PresenterWireFormatType.CPNav) ||
				(format == PresenterWireFormatType.CPCapability))
			{
				result = ProcessCP();
			}
            else if (format == PresenterWireFormatType.CP3) {
                result =  ProcessCP3();
            }
            else if (format == PresenterWireFormatType.RTDocument) {
                result =  ProcessRTDoc();
            }
            else if (format == PresenterWireFormatType.Video) {
                if (this.presFromVideoMgr != null) {
                    result = this.presFromVideoMgr.Process();
                }
            }
            if ((dataList != null) && (dataList.Count != 0)) {
                dataList.Sort();
            }
			return result;
		}


		/// <summary>
		/// Cancel the Process method if it is running.
		/// </summary>
		public void Stop()
		{
			stopnow = true;
            if (this.presFromVideoMgr != null) {
                this.presFromVideoMgr.Stop();
            }
		}

		/// <summary>
		/// Emit our results to a string.
		/// </summary>
		/// <returns></returns>
		public String Print()
		{
            if (this.presFromVideoMgr != null) {
                return this.presFromVideoMgr.Print();
            }

			//Print slide titles followed by dataList.
			StringBuilder sb = new StringBuilder();
			sb.Append(slideTitleMonitor.Print());
            if ((format == PresenterWireFormatType.CP3) && (this.cp3Mgr != null)) {
                sb.Append(this.cp3Mgr.GetTitles());
            }
			foreach (DataItem di in dataList)
			{
				sb.Append(di.Print());
			}
			return sb.ToString();
		}

		#endregion Public Methods

		#region Public Static

		/// <summary>
		/// Examine the first frames of a stream. Return true if it is a valid presentation stream.
		/// If so, also indicate its wire format and role.
		/// </summary>
		/// <param name="streamID"></param>
		/// <returns></returns>
		public static bool GetPresentationType(Stream stream, out PresenterWireFormatType wireFormat, out PresenterRoleType role,
            out string presenterName)
		{
            presenterName = "";
			wireFormat = PresenterWireFormatType.Unknown;
			role = PresenterRoleType.Unknown;
			PayloadType payload = Utility.StringToPayloadType(stream.Payload);
			//Get 5 seconds of data
			DBStreamPlayer dbsp = new DBStreamPlayer(stream.StreamID, Constants.TicksPerSec*stream.Seconds, payload);
			BufferChunk bc;
			long timestamp;
			PresenterWireFormatType thisFormat; 
			PresenterRoleType thisRole; 
			while (dbsp.GetNextFrame(out bc,out timestamp))
			{
				AnalyzeFrame(bc, payload, out thisFormat, out thisRole, out presenterName);		
				if (thisFormat != PresenterWireFormatType.Unknown)
					wireFormat = thisFormat;
				if (thisRole != PresenterRoleType.Unknown)
					role = thisRole;
                if ((wireFormat != PresenterWireFormatType.Unknown) &&
                    (role != PresenterRoleType.Unknown)) {
                        return true;
                }
			}

			return false;
		}

		#endregion Public Static

		#region Private Methods

		/// <summary>
		/// Read CP data and process it.  Return error message if any, or null if successful.
		/// </summary>
		/// <returns></returns>
		/// Terminate asap if stopnow becomes true.
		private String ProcessCP()
		{
			BufferChunk frame;
			long timestamp;

			//Note: there would be multiple streams here if the presenter left the venue and rejoined during the segment.
			// This is in the context of a single segment, so we can assume one cname, and no overlapping streams.
			int[] streams = DatabaseUtility.GetStreams(payload,cname,null,start,end);

			if (streams==null)
			{
				Debug.WriteLine("No presentation data found.");
				return "Warning: No presentation data was found for the given time range for " + cname + " with PresentationFormat " +
					format.ToString();
			}

			DBStreamPlayer[] streamPlayers = new DBStreamPlayer[streams.Length];
			for(int i=0; i<streams.Length; i++)
			{
				streamPlayers[i] = new DBStreamPlayer(streams[i],start,end,payload);
			}

			//PRI3: it would be a good enhancement to do a look-back to establish any existing presentation state at the 
			// time of the beginning of the segment.  The slide should be quick to establish, but any ink that is 
			// already on slides will be missing.

			for (int i=0;i<streamPlayers.Length;i++)
			{
				while ((streamPlayers[i].GetNextFrame(out frame,out timestamp))&&(!stopnow))
				{
					switch (format)
					{
						case PresenterWireFormatType.CPNav:
						{
							ProcessCPNavFrame(frame,timestamp-offset);
							break;
						}
						case PresenterWireFormatType.CPCapability:
						{
							ProcessCPCapabilityFrame(frame,timestamp-offset);
							break;
						}
						default:
							break;
					}
				}
				if (stopnow)
					break;
			}
			return null;
		}

        /// <summary>
        /// Read CP data and process it.  Return error message if any, or null if successful.
        /// </summary>
        /// <returns></returns>
        /// Terminate asap if stopnow becomes true.
        private String ProcessCP3() {
            BufferChunk frame;
            long timestamp;
            long lookbehindto = this.m_ConferenceStart;
            if (m_PreviousSegmentEnd != 0) {
                lookbehindto = m_PreviousSegmentEnd;
            }
            //Note: there would be multiple streams here if the presenter left the venue and rejoined during the segment.
            // This is in the context of a single segment, so we can assume one cname, and no overlapping streams.
            // Get streams from the beginning of the whole conference so that we can get existing presentation state.
            int[] streams = DatabaseUtility.GetStreams(payload, cname, null, lookbehindto, end);

            if (streams == null) {
                Debug.WriteLine("No presentation data found.");
                return "Warning: No presentation data was found for the given time range for " + cname + " with PresentationFormat " +
                    format.ToString();
            }

            DBStreamPlayer[] streamPlayers = new DBStreamPlayer[streams.Length];
            for (int i = 0; i < streams.Length; i++) {
                streamPlayers[i] = new DBStreamPlayer(streams[i], lookbehindto, end, payload);
            }

            cp3Mgr = new CP3Manager.CP3Manager();
            initialState = new List<object>();
            rtUpdate = null;

            for (int i = 0; i < streamPlayers.Length; i++) {
                while ((streamPlayers[i].GetNextFrame(out frame, out timestamp)) && (!stopnow)) {
                    ProcessCP3Frame(cp3Mgr, frame, timestamp);
                }
                
                if (stopnow)
                    break;
            }

            //This covers an edge case that may occur if there are no slide transitions during the selected
            // part of the archive:
            if ((this.rtUpdate != null) && (dataList.Count == 0)) {
                dataList.Add(new DataItem(this.start - this.offset, CopyRTUpdate(rtUpdate)));
                rtUpdate = null;          
            }

            if (cp3Mgr.WarningLog != null) {
                foreach (string s in cp3Mgr.WarningLog) {
                    errorLog.WriteLine(s);
                }
                errorLog.ErrorLevel = 5;
            }

            return null;
        }

		/// <summary>
		/// Read RTDoc data and process it.  Return error message if any, or null if successful.
		/// </summary>
		/// <returns></returns>
		/// Terminate asap if stopnow becomes true.
		/// The CXP Presenter tool (and possibly other capabilities) use a regular sender and a background sender.
		/// This results in two RTP Streams being opened.  All Page messages (and possibly other messages)
		/// are sent on the background stream, while at least navigation and ink messages are sent on the other.  The
		/// streams run concurrently.  To handle this correctly we need to get all the streams from the given
		/// cname for the entire conference, then we might need to interleave the messages in the proper order.
		private String ProcessRTDoc()
		{
			Debug.Assert(format==PresenterWireFormatType.RTDocument);

			BufferChunk frame;
			long timestamp;
			firstRTDocFrame=true;

			//cache the RTDocuments and stray pages we find in the streams.
			rtDocs = new Hashtable();
			unresolvedPages = new Hashtable();

			//Get the start time for the entire conference and use that to get streams.
			long confStart = DatabaseUtility.GetConferenceStartTime(payload,cname,start,end);
			if (confStart <= 0)
			{
				Debug.WriteLine("No conference matches the presentation descriptor.");
				return "Warning: No conference exists in the database that matches this presentation: " + cname + " with PresentationFormat " +
					format.ToString();
			}

			int[] streams = DatabaseUtility.GetStreams(payload,cname,null,confStart,end);

			if (streams==null)
			{
				Debug.WriteLine("No presentation data found.");
				return "Warning: No presentation data was found for the given time range for " + cname + " with PresentationFormat " +
					format.ToString();
			}

			//For RTDocuments we need to run through the streams from the beginning to find any RTDocument and 
			// Page messages that were sent before the scope of the archive.  If any were sent before the 
			// archiver started, we're out of luck.
			DBStreamPlayer[] streamPlayers = new DBStreamPlayer[streams.Length];
			for(int i=0; i<streams.Length; i++)
			{
				streamPlayers[i] = new DBStreamPlayer(streams[i],0,end,payload);
			}

			//Interleave the streams since they may be concurrent.
			bool done = false;
			while((!stopnow) && (!done))
			{
				int minIndex = -1;
				long minTime = long.MaxValue;
				long nextTime;
				for (int i=0;i<streamPlayers.Length;i++)
				{
					if (streamPlayers[i].GetNextFrameTime(out nextTime))
					{
						if (nextTime<minTime)
						{
							minTime = nextTime;
							minIndex = i;
						}
					}
				}
				if (minIndex == -1)
				{
					break;
				}

				if ((streamPlayers[minIndex].GetNextFrame(out frame,out timestamp))&&(!stopnow))
				{
					ProcessRTDocFrame(frame,timestamp-offset);
				}
			}


			if (!stopnow)
			{
				//If there were any screenshot or misc images, add the RTDocs to the collection:
				if (this.miscRtDoc != null)
				{
					this.rtDocs.Add(this.miscRtDoc.Identifier,this.miscRtDoc);
				}
			}
			return null;
		}

		/// <summary>
		/// Process a frame in a RTDocument stream.  Note the timestamp may preceed the beginning of the archive,
		/// in which case we just cache RTDocuments and pages, but write no output.
		/// </summary>
		/// <param name="frame"></param>
		/// <param name="timestamp"></param>
		private void ProcessRTDocFrame(BufferChunk frame, long timestamp)
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

			if (timestamp<start)
			{
				//Just review the frame, accumulating RTDocuments and images
				if (rtobj is RTDocument)
				{
					addRtDocument((RTDocument)rtobj);
				}
				else if (rtobj is Page)
				{
					addPage((Page)rtobj);
				}
				else if (rtobj is RTPageAdd)
				{
					addRtPageAdd((RTPageAdd)rtobj);
				}
				else if (rtobj is RTNodeChanged)
				{
					//also track most recent nav so that we have a better chance of having correct state at "time zero"
					receiveRTNodeChanged((RTNodeChanged)rtobj,timestamp, false);
				}
				//PRI2: could also track ink so that ink state could be (more) correct at time zero.
			}
			else
			{
				if (firstRTDocFrame)
				{
					FilterRTUpdate(start);
					firstRTDocFrame = false;
				}

				//These frames may also cause nav data to be written to the output.
				if (rtobj is RTDocument)
				{
					addRtDocument((RTDocument)rtobj);
					//There is an implicit navigation to the first slide here
					UpdateRTUpdate(((RTDocument)rtobj).Identifier, 0, DeckTypeEnum.Presentation);
					FilterRTUpdate(timestamp);		

				}
				else if (rtobj is Page)
				{
					//These are slide deck pages
					addPage((Page)rtobj);
				}
				else if (rtobj is RTPageAdd)
				{
					//These are dynamically added pages such as WB and screenshots
					addRtPageAdd((RTPageAdd)rtobj);
				}
				else if (rtobj is RTNodeChanged)
				{
					receiveRTNodeChanged((RTNodeChanged)rtobj,timestamp,true);
				}
				else if (rtobj is RTStroke)
				{
					receiveRTStroke((RTStroke)rtobj,timestamp);
				}
				else if (rtobj is RTStrokeRemove)
				{
					receiveRTStrokeRemove((RTStrokeRemove)rtobj, timestamp);
				}
				else if (rtobj is RTFrame)
				{
					RTFrame rtf = (RTFrame)rtobj;
					if (rtf.ObjectTypeIdentifier == Constants.RTDocEraseAllGuid)
                    {
						processEraseLayer(timestamp);
                    }
                    else
                    {
                        Debug.WriteLine("Unhandled RTFrame type.");
                    }
				}
				else
				{
					Debug.WriteLine("Unhandled RT obj:" + rtobj.ToString());
				}
			}
		}


		private void receiveRTStrokeRemove(RTStrokeRemove rtsr, long timestamp)
		{
			ArchiveRTNav.RTDeleteStroke rtds;
			int pageIndex;

			if (rtDocs.ContainsKey(rtsr.DocumentIdentifier))
			{
				RTDocument rtd = (RTDocument)rtDocs[rtsr.DocumentIdentifier];
				pageIndex = RTDocumentUtility.PageIDToPageIndex(rtd,rtsr.PageIdentifier);
				if (pageIndex >= 0)
				{
					rtds = new ArchiveRTNav.RTDeleteStroke(rtsr.StrokeIdentifier,rtd.Identifier,pageIndex);
					dataList.Add(new DataItem(timestamp,rtds));
					return;
				}
			}

			if (this.miscRtDoc != null)
			{
				if (miscRtDoc.Resources.Pages.ContainsKey(rtsr.PageIdentifier))
				{
					pageIndex = RTDocumentUtility.PageIDToPageIndex(miscRtDoc,rtsr.PageIdentifier);
					if (pageIndex >= 0)
					{
						rtds = new ArchiveRTNav.RTDeleteStroke(rtsr.StrokeIdentifier,miscRtDoc.Identifier,pageIndex);
						dataList.Add(new DataItem(timestamp,rtds));
					}	
					else
					{
						Debug.WriteLine("Bad page index!!");
					}
					return;
				}
			}


			//Otherwise we assume the stroke is on a whiteboard.
			bool pageExists = false;
			if (this.wbRtDoc != null)
			{
				if (wbRtDoc.Resources.Pages.ContainsKey(rtsr.PageIdentifier))
				{
					pageExists = true;
				}
			}

			if (!pageExists)
			{
				//Make a new WB page with this identifier
				this.addToWbRtDoc(rtsr);
				//I believe there is also an implicit navigation here.
				UpdateRTUpdate(this.wbRtDoc.Identifier, RTDocumentUtility.PageIDToPageIndex(wbRtDoc,rtsr.PageIdentifier), 
					DeckTypeEnum.Whiteboard);
				FilterRTUpdate(timestamp);		
			}

			pageIndex = RTDocumentUtility.PageIDToPageIndex(wbRtDoc,rtsr.PageIdentifier);
			rtds = new ArchiveRTNav.RTDeleteStroke(rtsr.StrokeIdentifier,wbRtDoc.Identifier,pageIndex);
			dataList.Add(new DataItem(timestamp,rtds));
		}

		private void receiveRTStroke(RTStroke rts, long timestamp)
		{
			ArchiveRTNav.RTDrawStroke rtds;
			int pageIndex;
			Microsoft.Ink.Ink ink;

			if (rtDocs.ContainsKey(rts.DocumentIdentifier))
			{
				RTDocument rtd = (RTDocument)rtDocs[rts.DocumentIdentifier];
				pageIndex = RTDocumentUtility.PageIDToPageIndex(rtd,rts.PageIdentifier);
				if (pageIndex >= 0)
				{
					ink = rts.Stroke.Ink.Clone();
					for (int i = 0; i < ink.Strokes.Count; i++)
						ink.Strokes[i].Scale(500f / 960f, 500f / 720f);  // is this right?

					rtds = new ArchiveRTNav.RTDrawStroke(ink,rts.StrokeIdentifier,
						true,rts.DocumentIdentifier,pageIndex);
					dataList.Add(new DataItem(timestamp,rtds));
					return;
				}
			}

			if (this.miscRtDoc != null)
			{
				if (miscRtDoc.Resources.Pages.ContainsKey(rts.PageIdentifier))
				{
					pageIndex = RTDocumentUtility.PageIDToPageIndex(miscRtDoc,rts.PageIdentifier);
					if (pageIndex >= 0)
					{
						//These screenshots frequently have non-standard aspect ratio.
						int w = miscRtDoc.Resources.Pages[rts.PageIdentifier].Image.Width;
						int h = miscRtDoc.Resources.Pages[rts.PageIdentifier].Image.Height;
						float wscale = 960f;
						float hscale = 720f;
						int wclip=10000;
						int hclip=10000;
						double ar = ((double)w)/((double)h);
						if (ar > 1.34) //wide image
						{
							hscale = 720f * (float)(ar/1.333);
							hclip = (h*10000/w);
						}
						else if (ar < 1.32) //tall image
						{
							wscale = 960f/((float)(ar/1.333)); 
							wclip = (w*10000/h);
						}

						ink = rts.Stroke.Ink.Clone();
						for (int i = 0; i < ink.Strokes.Count; i++)
						{
							ink.Strokes[i].Scale(500f / wscale, 500f / hscale); 
						}
						
						//PRI2: still some work to do on clipping.
						//start by calculating the maximum w and h for ink in the coordinate space, then
						//for example if the bounding box x+w exceeds the max width, clip it.
						//I think I need to do an additional transform on the ink coordinate space to match
						//the bitmap coordinates.

						rtds = new ArchiveRTNav.RTDrawStroke(ink,rts.StrokeIdentifier,
							true,miscRtDoc.Identifier,pageIndex);
						dataList.Add(new DataItem(timestamp,rtds));
					}	
					else
					{
						Debug.WriteLine("Bad page index!!");
					}
					return;
				}
			}

			//Otherwise we assume the stroke is on a whiteboard.
			bool pageExists = false;
			
			if (this.wbRtDoc != null)
			{
				if (wbRtDoc.Resources.Pages.ContainsKey(rts.PageIdentifier))
				{
					pageExists = true;
				}
			}

			if (!pageExists)
			{
				//Make a new WB page with this identifier
				this.addToWbRtDoc(rts);
				//PRI1: I believe there is also an implicit navigation here.
				UpdateRTUpdate(this.wbRtDoc.Identifier, RTDocumentUtility.PageIDToPageIndex(wbRtDoc,rts.PageIdentifier), 
					DeckTypeEnum.Whiteboard);
				FilterRTUpdate(timestamp);		
			}

			pageIndex = RTDocumentUtility.PageIDToPageIndex(wbRtDoc,rts.PageIdentifier);

			ink = rts.Stroke.Ink.Clone();
			for (int i = 0; i < ink.Strokes.Count; i++)
				ink.Strokes[i].Scale(500f / 960f, 500f / 720f);  // is this right?

			rtds = new ArchiveRTNav.RTDrawStroke(ink,rts.StrokeIdentifier,
				true,wbRtDoc.Identifier,pageIndex);
			dataList.Add(new DataItem(timestamp,rtds));
		}

		private void receiveRTNodeChanged(RTNodeChanged rtnc, long timestamp, bool sendMessages)
		{
			foreach(Guid g in rtDocs.Keys)
			{
				RTDocument rtd = (RTDocument)rtDocs[g];
				if (rtd.Organization.TableOfContents.ContainsKey(rtnc.OrganizationNodeIdentifier))
				{
					UpdateRTUpdate(rtd.Identifier,
						rtd.Organization.TableOfContents.IndexOf(rtd.Organization.TableOfContents[rtnc.OrganizationNodeIdentifier]),
						DeckTypeEnum.Presentation);
					if (sendMessages)
						FilterRTUpdate(timestamp);		
					return;
				}
			}

			if (this.miscRtDoc != null)
			{
				if (miscRtDoc.Organization.TableOfContents.ContainsKey(rtnc.OrganizationNodeIdentifier))
				{
					UpdateRTUpdate(miscRtDoc.Identifier,
						miscRtDoc.Organization.TableOfContents.IndexOf(miscRtDoc.Organization.TableOfContents[rtnc.OrganizationNodeIdentifier]),
						DeckTypeEnum.Presentation);
					if (sendMessages)
						FilterRTUpdate(timestamp);		
					return;
				}
			}

			//If we get here we assume a navigation to a WB page.
			bool wbPageExists = false;
			if (this.wbRtDoc != null)
			{
				if (wbRtDoc.Organization.TableOfContents.ContainsKey(rtnc.OrganizationNodeIdentifier))
				{
					wbPageExists = true;
				}
			}

			if (!wbPageExists)
			{
				//Make a new TOC entry in the WB doc
				this.addTocIdToWbRtDoc(rtnc.OrganizationNodeIdentifier);
			}

			UpdateRTUpdate(wbRtDoc.Identifier,
				wbRtDoc.Organization.TableOfContents.IndexOf(wbRtDoc.Organization.TableOfContents[rtnc.OrganizationNodeIdentifier]),
				DeckTypeEnum.Whiteboard);
			if (sendMessages)
				FilterRTUpdate(timestamp);		

		}

		private void addRtDocument(RTDocument rtd)
		{
			if (!this.rtDocs.ContainsKey(rtd.Identifier))
			{
				rtDocs.Add(rtd.Identifier,rtd);
				slideTitleMonitor.ReceiveRTDocument(rtd);
			}
		}

		private void addPage(Page p)
		{
			//If it belongs to a deck we know about, add the page to the RTDocument
			foreach (Guid g in rtDocs.Keys)
			{
				if (RTDocumentUtility.AddPageToRTDocument((RTDocument)rtDocs[g],p))
					return;
			}

			//Otherwise, just store in unresolvedPages to process later. 
			//This could happen if the archiving started while a deck was being transmitted.. maybe this is not
			//a supported scenario anyway.
			if (!unresolvedPages.ContainsKey(p.Identifier))
			{
				unresolvedPages.Add(p.Identifier,p);
			}
			else
			{
				Debug.WriteLine("Warning: Duplicate page ID in unresolvedPages");
			}
		}

		private void addRtPageAdd(RTPageAdd rtpa)
		{
			if (rtpa.Page.Image == null)
			{
				addToWbRtDoc(rtpa);
			}
			else
			{
				addToMiscRtDoc(rtpa);
			}
		}

		private void addToWbRtDoc(RTPageAdd rtpa)
		{
			if (wbRtDoc == null)
			{
				wbRtDoc = new RTDocument();
				wbRtDoc.Identifier = Guid.NewGuid();
			}
			RTDocumentUtility.AddRTPageAddToRTDocument(wbRtDoc,rtpa);
		}

		private void addToWbRtDoc(RTStrokeRemove rtsr)
		{
			if (wbRtDoc == null)
			{
				wbRtDoc = new RTDocument();
				wbRtDoc.Identifier = rtsr.DocumentIdentifier;
			}
			RTDocumentUtility.AddPageIDToRTDocument(wbRtDoc, rtsr.PageIdentifier);
		}

		private void addToWbRtDoc(RTStroke rts)
		{
			if (wbRtDoc == null)
			{
				wbRtDoc = new RTDocument();
				wbRtDoc.Identifier = rts.DocumentIdentifier;
			}
			RTDocumentUtility.AddPageIDToRTDocument(wbRtDoc,rts.PageIdentifier);
		}

		private void addTocIdToWbRtDoc(Guid tocID)
		{
			if (wbRtDoc == null)
			{
				wbRtDoc = new RTDocument();
				wbRtDoc.Identifier = Guid.NewGuid();
			}
			RTDocumentUtility.AddTocIDToRTDocument(wbRtDoc,tocID);
			
		}

		private void addToMiscRtDoc(RTPageAdd rtpa)
		{
			if (this.miscRtDoc == null)
			{
				this.miscRtDoc = new RTDocument();
				miscRtDoc.Identifier = Guid.NewGuid();
			}
			RTDocumentUtility.AddRTPageAddToRTDocument(this.miscRtDoc,rtpa);
			slideTitleMonitor.ReceiveRTPageAdd(rtpa,miscRtDoc.Identifier);

		}


		/// <summary>
		/// Examine one presentation frame and adjust data structures and write to output as appropriate.
		/// In this case we already know that the stream was produced by Classroom Presenter when running
		/// as a CXP Capability or in "RTDocuments" mode.
		/// </summary>
		/// When used as a capability or in "RTDocuments" mode, Classroom Presenter sends some messages 
		/// in its native format, packages other messages in RTFrame objects, and puts still others in
		/// RTFrame objects which are in turn packed inside other RT objects.
		/// <param name="frame"></param>
		/// <param name="timestamp"></param>
		private void ProcessCPCapabilityFrame(BufferChunk frame, long timestamp)
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

			// RTNodeChanged , RTStrokeAdd, and RTStrokeRemove contain a RTFrame which contains the CP message.
			if (rtobj is RTNodeChanged) 
			{
				//Classroom Presenter wraps CPPageUpdate inside the RTFrame which is in turn wrapped in
				// the RTNodeChanged Extension.
				if (((RTNodeChanged)rtobj).Extension is RTFrame)
				{
					RTFrame rtf = (RTFrame)((RTNodeChanged)rtobj).Extension;
					if (rtf.ObjectTypeIdentifier==Constants.PageUpdateIdentifier)
					{
						if (rtf.Object is PresenterNav.CPPageUpdate)
						{
							UpdateRTUpdate((PresenterNav.CPPageUpdate)rtf.Object);
							FilterRTUpdate(timestamp);							
						}
					}
				}
			}
			else if (rtobj is RTStroke)
			{
				if (((RTStroke)rtobj).Extension is RTFrame)
				{
					RTFrame rtf = (RTFrame)((RTStroke)rtobj).Extension;
					if (rtf.ObjectTypeIdentifier == Constants.StrokeDrawIdentifier)
					{
						if (rtf.Object is PresenterNav.CPDrawStroke)
						{
							PresenterNav.CPDrawStroke cpds = (PresenterNav.CPDrawStroke)rtf.Object;
							ArchiveRTNav.RTDrawStroke rtds = new ArchiveRTNav.RTDrawStroke(cpds.Stroke.Ink,cpds.Stroke.Guid,cpds.Stroke.StrokeFinished,cpds.DocRef.GUID,cpds.PageRef.Index);
							dataList.Add(new DataItem(timestamp,rtds));
						}
					}
				}
			}
			else if (rtobj is RTStrokeAdd)
			{
				if (((RTStrokeAdd)rtobj).Extension is RTFrame)
				{
					RTFrame rtf = (RTFrame)((RTStrokeAdd)rtobj).Extension;
					if (rtf.ObjectTypeIdentifier == Constants.StrokeDrawIdentifier)
					{
						if (rtf.Object is PresenterNav.CPDrawStroke)
						{
							PresenterNav.CPDrawStroke cpds = (PresenterNav.CPDrawStroke)rtf.Object;
							ArchiveRTNav.RTDrawStroke rtds = new ArchiveRTNav.RTDrawStroke(cpds.Stroke.Ink,cpds.Stroke.Guid,cpds.Stroke.StrokeFinished,cpds.DocRef.GUID,cpds.PageRef.Index);
							dataList.Add(new DataItem(timestamp,rtds));
						}
					}
				}
			}
			else if (rtobj is RTStrokeRemove)
			{
				if (((RTStrokeRemove)rtobj).Extension is RTFrame)
				{
					RTFrame rtf = (RTFrame)((RTStrokeRemove)rtobj).Extension;
					if (rtf.ObjectTypeIdentifier == Constants.StrokeDeleteIdentifier)
					{
						if (rtf.Object is PresenterNav.CPDeleteStroke)
						{
							ArchiveRTNav.RTDeleteStroke rtds = new ArchiveRTNav.RTDeleteStroke(((PresenterNav.CPDeleteStroke)rtf.Object).Guid,
								((PresenterNav.CPDeleteStroke)rtf.Object).DocRef.GUID,
								((PresenterNav.CPDeleteStroke)rtf.Object).PageRef.Index);
							dataList.Add(new DataItem(timestamp,rtds));
						}
					}
				}
			}
				//Other CP messages are sent directly in RTFrame
			else if (rtobj is RTFrame)
			{
				if (((RTFrame)rtobj).ObjectTypeIdentifier==Constants.CLASSROOM_PRESENTER_MESSAGE)
				{
					ProcessCPMessage(((RTFrame)rtobj).Object,timestamp);
				}
				else if (((RTFrame)rtobj).ObjectTypeIdentifier == Constants.RTDocEraseAllGuid)
				{
					//This case is a bit ugly because the RTDoc message contains no info about the deck and slide.
					// We need to apply to the current deck and slide (if any)
					processEraseLayer(timestamp);
				}

			}
				//Still other messages are in native format
			else if (rtobj is PresenterNav.CPMessage)
			{
				//CP still sends some messages in native format
				ProcessCPMessage(rtobj,timestamp);
			}
		}

		/// <summary>
		/// Examine one presentation frame and adjust data structures and write to output as appropriate.
		/// In this case we know the stream was produced by Classroom Presenter running as a stand-alone app, so we
		/// only look for the native CP messages.
		/// </summary>
		/// <param name="frame"></param>
		/// <param name="timestamp"></param>
		private void ProcessCPNavFrame(BufferChunk frame, long timestamp)
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
			
			ProcessCPMessage(rtobj,timestamp);
		}

		private void ProcessCP3Frame(CP3Manager.CP3Manager cp3Mgr, BufferChunk frame, long timestamp)
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
				Debug.WriteLine("ProcessCP3Frame exception deserializing message. size=" + frame.Length.ToString() +
					" exception=" + e.ToString());
				return;
			}
			
            List<object> archiveObjects = cp3Mgr.Accept(rtobj);

            if ((archiveObjects != null) && (archiveObjects.Count > 0)) {
                if (timestamp >= this.start) {
                    if (initialState.Count > 0) { 
                        //insert messages to establish initial presentation state
                        //Filter out non-terminal text annotations
                        initialState = filterNonTerminalTextAnnotations(initialState);
                        //Filter out deleteStrokes that have no corresponding addStroke.
                        initialState = filterDeleteStrokes(initialState);
                        long counter = 0;
                        foreach (object rto in initialState) {
                            //Note: we use this counter because the DataItems will be sorted by timestamp later, and we need to
                            //preserve the order of RTObjects.  E.g. Reordering draw and delete causes stray ink.
                            dataList.Add(new DataItem(this.start + counter - offset, rto));
                            counter++;
                            ///Debugging code:
                            //DateTime dt = new DateTime(this.start - offset);
                            //if (rto is RTUpdate) {
                            //    Debug.WriteLine("RTUpdate at " + dt.ToString("M/d/yyyy HH:mm:ss.ff"));
                            //}
                            //if (rto is RTDrawStroke) {
                            //    RTDrawStroke rtds = (RTDrawStroke)rto;
                            //    Debug.WriteLine("InitialState RTDrawStroke at " + dt.ToString("M/d/yyyy HH:mm:ss.ff") + "; slide=" + rtds.SlideIndex.ToString() + "; deck=" + rtds.DeckGuid.ToString());
                            //}
                            //else if (rto is RTDeleteStroke) {
                            //    RTDeleteStroke rtds = (RTDeleteStroke)rto;
                            //    Debug.WriteLine("InitialState RTDeleteStroke at " + dt.ToString("M/d/yyyy HH:mm:ss.ff") + "; slide=" + rtds.SlideIndex.ToString() + "; deck=" + rtds.DeckGuid.ToString());
                            //}
                        }
                        initialState.Clear();
                    }
                    if (rtUpdate != null) {
                        dataList.Add(new DataItem(this.start - offset, CopyRTUpdate(rtUpdate)));
                        rtUpdate = null;
                    }

                    foreach (object o in archiveObjects) {
                        if (o is RTUpdate) {
                            DateTime dt = new DateTime(timestamp - offset);
                            Debug.WriteLine("RTUpdate at " + dt.ToString("M/d/yyyy HH:mm:ss.ff"));
                        }
                        DataItem di = new DataItem(timestamp - offset, o);
                        Debug.WriteLine(di.ToString());
                        dataList.Add(di);
                    }
                }
                else {
                    //Remember messages needed to establish initial state
                    foreach (object rto in archiveObjects) {
                        if (rto is ArchiveRTNav.RTDeleteStroke) 
                        {
                            initialState.Add(rto);
                        }
                        else if (rto is ArchiveRTNav.RTDrawStroke) {
                            RTDrawStroke rtds = (RTDrawStroke)rto;
                            //Ignore real-time strokes.
                            if (rtds.StrokeFinished) {
                                initialState.Add(rto);
                            }
                        }
                        else if ((rto is ArchiveRTNav.RTTextAnnotation) || 
                                (rto is ArchiveRTNav.RTQuickPoll) ||
                                (rto is ArchiveRTNav.RTImageAnnotation)) {
                            initialState.Add(rto);
                        }
                        else if ((rto is RTEraseLayer) || (rto is RTEraseAllLayers) || 
                            (rto is RTDeleteTextAnnotation) || (rto is RTDeleteAnnotation)) {
                            initialState.Add(rto);
                        }
                        else if (rto is ArchiveRTNav.RTUpdate) {
                            //We only need the most recent navigation message.
                            this.rtUpdate = (ArchiveRTNav.RTUpdate)rto;
                        }
                    }

                }
            }
		}

        private List<object> filterNonTerminalTextAnnotations(List<object> initialState) {
            Hashtable taIds = new Hashtable();
            List<object> newInitialState = new List<object>();
            for (int i = initialState.Count - 1; i >= 0; i--) {
                object o = initialState[i];
                if (o is RTTextAnnotation) {
                    RTTextAnnotation rtta = (RTTextAnnotation)o;
                    if (taIds.ContainsKey(rtta.Guid)) {
                        continue;
                    }
                    taIds.Add(rtta.Guid, null);
                }
                newInitialState.Add(o);
            }
            newInitialState.Reverse();
            return newInitialState;
        }

        /// <summary>
        /// Filter out the deletes for which there is no corresponding stroke.
        /// </summary>
        /// <param name="initialState"></param>
        /// <returns></returns>
        private List<object> filterDeleteStrokes(List<object> initialState) {
            List<object> newInitialState = new List<object>();
            Hashtable strokes = new Hashtable();
            foreach (object o in initialState) {
                if (o is RTDrawStroke) {
                    RTDrawStroke rtds = (RTDrawStroke)o;
                    if (!strokes.ContainsKey(rtds.Guid))
                        strokes.Add(rtds.Guid, null);
                } else if (o is RTDeleteStroke) {
                    RTDeleteStroke rtds = (RTDeleteStroke)o;
                    if (!strokes.ContainsKey(rtds.Guid)) {
                        continue;
                    }
                }
                newInitialState.Add(o);
            }
            return newInitialState;
        }



		/// <summary>
		/// Handle all the important CP native messages.
		/// </summary>
		/// <param name="rtobj"></param>
		/// <param name="timestamp"></param>
		private void ProcessCPMessage(object rtobj, long timestamp)
		{
			DateTime thisdt = new DateTime(timestamp);
			String thisdts = thisdt.ToString();
			if (rtobj is PresenterNav.CPPageUpdate)
			{
				/// Presenter sends these once per second. These indicate current slide index, deck index
				/// and deck type.
				UpdateRTUpdate((PresenterNav.CPPageUpdate)rtobj);
				FilterRTUpdate(timestamp);
				//DumpPageUpdate((PresenterNav.CPPageUpdate)rtobj);
			}
			else if (rtobj is PresenterNav.CPScrollLayer)
			{
				///Presenter sends these once per second, and during a scroll operation.
				PresenterNav.CPScrollLayer cpsl = (PresenterNav.CPScrollLayer)rtobj;
				scrollPosCache.Add(cpsl.DocRef.GUID,cpsl.PageRef.Index,cpsl.ScrollPosition,cpsl.ScrollExtent);
				FilterRTScollLayer(cpsl,timestamp);
				//DumpScrollLayer((PresenterNav.CPScrollLayer)rtobj);
			}
			else if (rtobj is PresenterNav.CPDeckCollection) 
			{
				//Presenter sends this once per second.  The only thing we want to get from 
				//this is the slide size (former ScreenConfiguration)
				UpdateRTUpdate((PresenterNav.CPDeckCollection)rtobj,timestamp);
				//DumpDeckCollection((PresenterNav.CPDeckCollection)rtobj);

				//look for slide titles for new decks.  From WMG SlideTitleMonitor
				slideTitleMonitor.ReceivePresenterNav((PresenterNav.CPDeckCollection)rtobj);
			}
			else if (rtobj is WorkSpace.BeaconPacket) //the original beacon
			{
				//Presenter sends this once per second.  The only thing we want to get from 
				//this is the background color
				UpdateRTUpdate((WorkSpace.BeaconPacket)rtobj);
				//DumpBeacon((BeaconPacket)rtobj);
			}
			else if (rtobj is PresenterNav.CPDrawStroke) //add a stroke
			{
				PresenterNav.CPDrawStroke cpds = (PresenterNav.CPDrawStroke)rtobj;
				ArchiveRTNav.RTDrawStroke rtds = new ArchiveRTNav.RTDrawStroke(cpds.Stroke.Ink,cpds.Stroke.Guid,cpds.Stroke.StrokeFinished,cpds.DocRef.GUID,cpds.PageRef.Index);
				//every ink operation goes straight to the output.				
				dataList.Add(new DataItem(timestamp,rtds));
			}
			else if (rtobj is PresenterNav.CPDeleteStroke) //delete one stroke
			{
				ArchiveRTNav.RTDeleteStroke rtds = new ArchiveRTNav.RTDeleteStroke(((PresenterNav.CPDeleteStroke)rtobj).Guid,((PresenterNav.CPDeleteStroke)rtobj).DocRef.GUID,((PresenterNav.CPDeleteStroke)rtobj).PageRef.Index);
				dataList.Add(new DataItem(timestamp,rtds));
			}
			else if (rtobj is PresenterNav.CPEraseLayer) //clear all strokes from one page
			{
				ArchiveRTNav.RTEraseLayer rtel = new ArchiveRTNav.RTEraseLayer(((PresenterNav.CPEraseLayer)rtobj).DocRef.GUID,((PresenterNav.CPEraseLayer)rtobj).PageRef.Index);
				dataList.Add(new DataItem(timestamp,rtel));
			}
			else if (rtobj is PresenterNav.CPEraseAllLayers) //clear all strokes from a deck
			{
				ArchiveRTNav.RTEraseAllLayers rteal = new ArchiveRTNav.RTEraseAllLayers(((PresenterNav.CPEraseAllLayers)rtobj).DocRef.GUID);
				dataList.Add(new DataItem(timestamp,rteal));
			}
			else if (rtobj is WorkSpace.OverlayMessage) //Student submission received
			{
				//Just cache the overlay messages here -- CPPageUpdate tells us when to display the student submission.
				studentSubmissionCache.ReceivePresenterNav((WorkSpace.OverlayMessage)rtobj);
			}
            else if (rtobj is WorkSpace.ScreenConfigurationMessage) 
            {
                // nothing to do here.
            }
            else if (rtobj is WorkSpace.QuickPollMessage)
            {
                QuickPollMessage qpm = (QuickPollMessage)rtobj;
				Debug.WriteLine("QuickPollMessage");
            }
            else if (rtobj is WorkSpace.SlideMessage)
            {
                SlideMessage sm = (SlideMessage)rtobj;
                //This is a special case to handle the Presenter 2.1 QuickPoll mechanism.  Notice that:
                // -This functionality is not supported when presentation is used to generate the video
                // -It is not supported when CP is a CXP capability
                if ((sm.DeckType == "StudentSubmission") && (sm.DocRef.GUID == Guid.Empty))
                { 
                    slideMessageCache.Add(sm);
                }
                //System.Drawing.Image i = sm.Slide.GetImageWithDefault(SlideViewer.SlideMode.Shared);
                //There are some cases where it is important to process this: Quick poll histogram is one.
				Debug.WriteLine("SlideMessage: " + sm.ToString());
            }
            else
			{
				Type t = rtobj.GetType();
				Debug.WriteLine("Unhandled Type:" + t.ToString());
			}

		}

		/// <summary>
		/// CPPageUpdate gives us slide index, deck Guid, and deck type.  Just track the current values.
		/// </summary>
		/// <param name="cppu"></param>
		private void UpdateRTUpdate(PresenterNav.CPPageUpdate cppu)
		{  
			rtUpdate.SlideIndex = cppu.PageRef.Index;
			rtUpdate.DeckGuid = cppu.DocRef.GUID;
			if (cppu.DeckType == "Presentation")
			{
				rtUpdate.DeckType = (Int32)DeckTypeEnum.Presentation;
			}
			else if (cppu.DeckType == "WhiteBoard")
			{
				rtUpdate.DeckType = (Int32)DeckTypeEnum.Whiteboard;
			}
			else if (cppu.DeckType == "StudentSubmission")
			{
				rtUpdate.DeckType = (Int32)DeckTypeEnum.StudentSubmission;
			}		
		}

		private void UpdateRTUpdate(Guid deckGuid, int slideIndex, DeckTypeEnum type)
		{
			rtUpdate.SlideIndex = slideIndex;
			rtUpdate.DeckGuid = deckGuid;
			rtUpdate.DeckType = (Int32)type;
		}

		/// <summary>
		/// CPDeckCollection tells us the slide size.  Write to the output if it changed.
		/// </summary>
		/// <param name="cpdc"></param>
		/// <param name="timestamp"></param>
		private void UpdateRTUpdate(PresenterNav.CPDeckCollection cpdc, long timestamp)
		{
			if (rtUpdate.SlideSize != cpdc.ViewPort.SlideSize)
			{
				rtUpdate.SlideSize = cpdc.ViewPort.SlideSize;
				if (rtUpdate.DeckGuid != Guid.Empty)
				{
					dataList.Add(new DataItem(timestamp,CopyRTUpdate(rtUpdate)));
				}
			}
		}

		/// <summary>
		/// The erase slide message in RTDocs contains no info about the deck and slide.  If there is a current deck
		/// and slide, apply erase to that, otherwise do nothing.
		/// </summary>
		private void processEraseLayer(long timestamp)
		{
			if ((this.lastDeckGuid != Guid.Empty) && (this.lastSlideIndex != -1))
			{
				ArchiveRTNav.RTEraseLayer rtel = new ArchiveRTNav.RTEraseLayer(lastDeckGuid,lastSlideIndex);
				dataList.Add(new DataItem(timestamp,rtel));
			}
		}

		/// <summary>
		/// BeaconPacket just tells us the background color.
		/// </summary>
		/// <param name="bp"></param>
		private void UpdateRTUpdate(BeaconPacket bp)
		{
			rtUpdate.BackgroundColor = bp.BGColor;
		}

		/// <summary>
		/// Emit a message to the output if a slide transition occurred since the last call.
		/// </summary>
		/// <param name="timestamp"></param>
		private void FilterRTUpdate(long timestamp)
		{
			if (rtUpdate.DeckGuid==Guid.Empty)
			{
				return;
			}

			///write the update message on transition.
			if ((lastSlideIndex != rtUpdate.SlideIndex) ||
				(lastDeckGuid != rtUpdate.DeckGuid))
			{
				
				//Set the scroll position for the new slide
				ScrollParams sp = scrollPosCache.Get(rtUpdate.DeckGuid,rtUpdate.SlideIndex);
				rtUpdate.ScrollPosition = sp.ScrollPos;
				rtUpdate.ScrollExtent = sp.ScrollExtent;

				if (rtUpdate.DeckType == (Int32)DeckTypeEnum.StudentSubmission)
				{
					//Output a Student Submission slide + possible overlay ink
					AddStudentSubmission(rtUpdate, timestamp);
                    DateTime t = new DateTime(timestamp);
                    Debug.WriteLine("student submission at " + t.ToString() +
                        ";deckassociation=" + rtUpdate.DeckAssociation.ToString() +
                        ";deckguid=" + rtUpdate.DeckGuid.ToString() +
                        ";slideassociation=" + rtUpdate.SlideAssociation.ToString() + 
                        ";slideindex=" + rtUpdate.SlideIndex.ToString());
                      
				}
				else
				{ 
					//Output a Whiteboard or presentation slide
					dataList.Add(new DataItem(timestamp,CopyRTUpdate(rtUpdate)));
				}

				lastSlideIndex=rtUpdate.SlideIndex;
				lastDeckGuid=rtUpdate.DeckGuid;			
			}
		}

        /// <summary>
        /// We have already determined that it is a student submission and there is no overlay message.
        /// If we have the corresponding slide message, we special case this to a Quick poll slide.  
        /// </summary>
        /// <param name="rtu"></param>
        /// <param name="timestamp"></param>
        private void AddQuickPoll(RTUpdate rtu, long timestamp)
        {
            if (slideMessageCache.ContainsKey(Guid.Empty, rtu.SlideIndex))
            {
                //Here we map to a presentation slide that is contained in SlideMessage.
                rtu.DeckAssociation = Guid.Empty;
                rtu.SlideAssociation = rtu.SlideIndex;

                dataList.Add(new DataItem(timestamp, CopyRTUpdate(rtu)));
            }

        }

		/// <summary>
		/// Output a Student Submission RTUpdate message, and possibly also the ink in its overlay.
		/// We keep a list of student submissions for which we have already sent the ink so we don't do it
		/// more than once.
		/// </summary>
		/// <param name="rtu"></param>
		private void AddStudentSubmission(RTUpdate rtu, long timestamp)
		{
			if (studentSubmissionCache == null)
				return;
            
			OverlayMessage om = studentSubmissionCache.Lookup(rtu.DeckGuid,rtu.SlideIndex);
            if (om == null)
            {
                AddQuickPoll(rtu, timestamp);
                return;

            }

			//Mapping to the presentation slide
			rtu.DeckAssociation = om.PresentationDeck.GUID;
			rtu.SlideAssociation = om.PDeckSlideIndex.Index;

			dataList.Add(new DataItem(timestamp,CopyRTUpdate(rtu)));

			//If this is the first time we've seen this slide, also output ink in overlay
			String key = rtu.DeckGuid.ToString() + "-" + rtu.SlideIndex.ToString();
			Guid guid;
			//This is the magic ID for Presenter ink
			Guid GUID_TAG = new Guid ("{179222d6-bcc1-4570-8d8f-7e8834c1dd2a}");
			if (! seenStudentSubmissions.ContainsKey(key))
			{
				seenStudentSubmissions.Add(key,true);

                if ((om.SlideOverlay.UserScribble != null) &&
                    (om.SlideOverlay.UserScribble is SlideViewer.InkScribble))
                {
                    Microsoft.Ink.Ink ink = ((SlideViewer.InkScribble)(om.SlideOverlay.UserScribble)).Ink;
                    Microsoft.Ink.Strokes strokes = ink.Strokes;

                    foreach (Microsoft.Ink.Stroke stroke in strokes)
                    {
                        guid = Guid.Empty;
                        ///Note that StudentSubmission ink may not have this extended property, which I guess doesn't matter since you never have to erase these strokes..
                        //Ignore any without a valid stroke guid
                        //if (stroke.ExtendedProperties.DoesPropertyExist(GUID_TAG))
                        //{
                        //    guid = new Guid((string)stroke.ExtendedProperties[GUID_TAG].Data);
                            Microsoft.Ink.Ink newInk = new Microsoft.Ink.Ink();
                            Microsoft.Ink.Strokes newStrokes = ink.CreateStrokes(new int[] { stroke.Id });
                            newInk.AddStrokesAtRectangle(newStrokes, newStrokes.GetBoundingBox());
                            ArchiveRTNav.RTDrawStroke rtds = new ArchiveRTNav.RTDrawStroke(newInk, guid, true, rtu.DeckGuid, rtu.SlideIndex);
                            dataList.Add(new DataItem(timestamp, rtds));
                        //}
                    }
                }

				if ((om.SlideOverlay.OtherScribble != null) &&  
					(om.SlideOverlay.OtherScribble is SlideViewer.InkScribble))
				{
					Microsoft.Ink.Ink ink = ((SlideViewer.InkScribble)(om.SlideOverlay.OtherScribble)).Ink;
					Microsoft.Ink.Strokes strokes = ink.Strokes;

					foreach (Microsoft.Ink.Stroke stroke in strokes)
					{
						guid = Guid.Empty;
						Microsoft.Ink.Ink newInk = new Microsoft.Ink.Ink();
						Microsoft.Ink.Strokes newStrokes = ink.CreateStrokes(new int[] { stroke.Id });
						newInk.AddStrokesAtRectangle(newStrokes, newStrokes.GetBoundingBox());
						ArchiveRTNav.RTDrawStroke rtds = new ArchiveRTNav.RTDrawStroke(newInk,guid,true,rtu.DeckGuid,rtu.SlideIndex);
						dataList.Add(new DataItem(timestamp,rtds));
					}
				}

			}
		}

		/// <summary>
		/// Send a scroll message if the scroll position changed since the last call.
		/// </summary>
		/// The scroll properties are per-slide, so a slide transition may result in a scroll 
		/// message being sent.  That actually may be a good thing.
		/// Note RTUpdate also carries periodic updates to scroll position.
		/// <param name="rtsl"></param>
		private void FilterRTScollLayer(PresenterNav.CPScrollLayer cpsl, long timestamp)
		{
			double scrollExtent = cpsl.ScrollExtent;
			double scrollPosition = cpsl.ScrollPosition;
			if	((scrollExtent != lastScrollExtent) ||
				(scrollPosition != lastScrollPosition))
			{
				lastScrollExtent = scrollExtent;
				lastScrollPosition = scrollPosition;
				ArchiveRTNav.RTScrollLayer rtsl = new ArchiveRTNav.RTScrollLayer(cpsl.ScrollPosition,cpsl.ScrollExtent,cpsl.DocRef.GUID,cpsl.PageRef.Index);
				dataList.Add(new DataItem(timestamp,rtsl));
			}
		}

		/// <summary>
		/// Make an identical copy of a RTUpdate object
		/// </summary>
		/// <param name="rtu"></param>
		/// <returns></returns>
		internal static RTUpdate CopyRTUpdate(RTUpdate rtu)
		{
			RTUpdate rtuOut = new RTUpdate();
			rtuOut.BackgroundColor = rtu.BackgroundColor;
			rtuOut.BaseUrl = rtu.BaseUrl;
			rtuOut.DeckAssociation = rtu.DeckAssociation;
			rtuOut.DeckGuid = rtu.DeckGuid;
			rtuOut.DeckType = rtu.DeckType;
			rtuOut.Extent = rtu.Extent;
			rtuOut.ScrollExtent = rtu.ScrollExtent;
			rtuOut.ScrollPosition = rtu.ScrollPosition;
			rtuOut.SlideAssociation = rtu.SlideAssociation;
			rtuOut.SlideIndex = rtu.SlideIndex;
			rtuOut.SlideSize = rtu.SlideSize;
			return rtuOut;
		}

		/// <summary>
		/// Examine a frame to try to determine wire format (eg. 'CPNav' or 'RTDocuments') and role (eg. Instructor, Shared display ...)
		/// </summary>
		/// <param name="frame"></param>
		/// <param name="wireFormat"></param>
		/// <param name="role"></param>
		private static void AnalyzeFrame(BufferChunk frame, PayloadType payload, 
			out PresenterWireFormatType wireFormat, out PresenterRoleType role, out string presenterName)
		{
			wireFormat = PresenterWireFormatType.Unknown;
			role = PresenterRoleType.Unknown;
            presenterName = "";

			//If the payload isn't one of the supported payloads, give up now.
			if ((payload != PayloadType.dynamicPresentation) &&
				(payload != PayloadType.RTDocument))
			{
				return;
			}

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

			if (payload == PayloadType.dynamicPresentation) 
			{
                if (rtobj is WorkSpace.BeaconPacket) {
                    //Classroom Presenter running Stand-alone.
                    wireFormat = PresenterWireFormatType.CPNav;
                    WorkSpace.BeaconPacket bp = (BeaconPacket)rtobj;
                    role = BeaconRoleToPresenterRoleType(bp.Role);
                }
                else if (rtobj is UW.ClassroomPresenter.Network.Chunking.Chunk) {
                    wireFormat = PresenterWireFormatType.CP3;
                    int r = CP3Manager.CP3Manager.AnalyzeChunk((UW.ClassroomPresenter.Network.Chunking.Chunk)rtobj, out presenterName);
                    if (r == 1)
                        role = PresenterRoleType.Instructor;
                    else if (r == 2)
                        role = PresenterRoleType.SharedDisplay;
                    else if (r == 3)
                        role = PresenterRoleType.Student;
                    else
                        role = PresenterRoleType.Unknown;
                }
			}	
			else if (payload == PayloadType.RTDocument) //CP as a capability, or the CXP presentation tool
			{
				if (rtobj is RTFrame)
				{
					RTFrame rtFrame = (RTFrame)rtobj;
					if (rtFrame.ObjectTypeIdentifier == Constants.CLASSROOM_PRESENTER_MESSAGE)
					{
						if (rtFrame.Object is WorkSpace.BeaconPacket)
						{
							//Classroom Presenter used as a capability
							WorkSpace.BeaconPacket bp = (BeaconPacket)rtFrame.Object;
							role = BeaconRoleToPresenterRoleType(bp.Role);
							wireFormat = PresenterWireFormatType.CPCapability;
						}
					}
				}
				else if (rtobj is RTNodeChanged) 
				{
					if (((RTNodeChanged)rtobj).Extension is RTFrame)
					{
						RTFrame rtf = (RTFrame)((RTNodeChanged)rtobj).Extension;
						if (rtf.ObjectTypeIdentifier==Constants.PageUpdateIdentifier)
						{
							if (rtf.Object is PresenterNav.CPPageUpdate)
							{
								role = PresenterRoleType.Instructor;
								wireFormat = PresenterWireFormatType.CPCapability;
							}
						}
					}
					else
					{
						role = PresenterRoleType.Instructor;
						wireFormat = PresenterWireFormatType.RTDocument;
					}
				}
				else if (rtobj is RTStroke)
				{
					if (((RTStroke)rtobj).Extension is RTFrame)
					{
						RTFrame rtf = (RTFrame)((RTStroke)rtobj).Extension;
						if (rtf.ObjectTypeIdentifier == Constants.StrokeDrawIdentifier)
						{
							if (rtf.Object is PresenterNav.CPDrawStroke)
							{
								//role = PresenterRoleType.Instructor;
								//wireFormat = PresenterWireFormatType.CPCapability;
							}
						}
					}
					else
					{
						role = PresenterRoleType.Instructor;
						wireFormat = PresenterWireFormatType.RTDocument;
					}
				}
				else if (rtobj is Page)
				{
					//CP NEVER sends these
					role = PresenterRoleType.Instructor;
					wireFormat = PresenterWireFormatType.RTDocument;
				}
			}
		}

		private static PresenterRoleType BeaconRoleToPresenterRoleType(WorkSpace.ViewerFormControl.RoleType bpRole)
		{
			PresenterRoleType role = PresenterRoleType.Other;

			if (bpRole == WorkSpace.ViewerFormControl.RoleType.Presenter)
			{
				role = PresenterRoleType.Instructor;
			}
			else if (bpRole == WorkSpace.ViewerFormControl.RoleType.SharedDisplay)
			{
				role = PresenterRoleType.SharedDisplay;
			}
			else if (bpRole == WorkSpace.ViewerFormControl.RoleType.Viewer)
			{
				role = PresenterRoleType.Student;
			}
			return role;
		}



		#endregion Private Methods

		#region DataItem Class

		/// <summary>
		/// Data representing a presentation frame to be archived.
		/// </summary>
		internal class DataItem : IComparable
		{
			private long timestamp;
			private object rtobj;
			private String type = "CXP3"; //Identifier for Presenter 2 (the only one we support here)

			public DataItem(long timestamp, object rtobj)
			{
				this.timestamp = timestamp;
				this.rtobj = rtobj;
			}

			/// <summary>
			/// Convert the frame to the XML element for storage.
			/// </summary>
			/// <returns></returns>
			public String Print()
			{
				DateTime dt = new DateTime(timestamp);
				BufferChunk bc = new BufferChunk(Utility.ObjectToByteArray(rtobj));
				String data = Convert.ToBase64String(bc.Buffer,bc.Index,bc.Length);

				return "<Script Time=\"" + dt.ToString("M/d/yyyy HH:mm:ss.ff") + "\" Type=\"" + type + 
					"\" Command=\"" + data + "\" />\r\n";
			}

            public override string ToString() {
				DateTime dt = new DateTime(timestamp);
                return "DataItem{time=" + dt.ToString("M/d/yyyy HH:mm:ss.ff") + ";type=" + rtobj.GetType().ToString() + "}";
            }

			public object RTObject
			{ get {return rtobj;}}

			public long Timestamp
			{ get {return timestamp;}}

            #region IComparable Members
            /// <summary>
            /// Support sorting by timestamp
            /// </summary>
            /// <param name="obj"></param>
            /// <returns></returns>
            public int CompareTo(object obj) {
                if (obj == null) return 1;
                return this.timestamp.CompareTo(((DataItem)obj).Timestamp);
            }

            #endregion
        }

		#endregion DataItem Class

		#region StudentSubmissionCache Class

		/// <summary>
		/// Store Ink, slide and deck references for student submissions (OverlayMessage).
		/// </summary>
		/// Support lookup by slide and deckGuid.
		private class StudentSubmissionCache
		{
			private Hashtable ssCache; //key is composed of deck Guid and slide index.  Value is the OverlayMessage

			public StudentSubmissionCache()
			{
				ssCache = new Hashtable();
			}

			public void ReceivePresenterNav(OverlayMessage om)
			{
				String key = om.FeedbackDeck.GUID.ToString() + "-" + om.SlideIndex.Index.ToString();
				lock (ssCache)
				{
					if (ssCache.ContainsKey(key))
						ssCache[key] = om;
					else
						ssCache.Add(key,om);
				}
			}

			public OverlayMessage Lookup(Guid deckGuid, Int32 slideIndex)
			{
				String key = deckGuid.ToString() + "-" + slideIndex.ToString();

				Debug.Assert(deckGuid != Guid.Empty);

				lock (ssCache)
				{
					if (ssCache.ContainsKey(key))
						return (OverlayMessage)ssCache[key];
				}

				return null;
			}
		}

		#endregion StudentSubmissionCache Class

        #region SlideMessageCache Class

        private class SlideMessageCache
        {
            private Hashtable smCache;

            public Hashtable Table
            {
                get { return smCache; }
            }

            public SlideMessageCache()
            {
                smCache = new Hashtable();
            }

            public void Add(SlideMessage sm)
            {
                String key = sm.DocRef.GUID.ToString() + "-" + sm.PageRef.Index.ToString();
                lock (smCache)
                {
                    if (smCache.ContainsKey(key))
                        smCache[key] = sm;
                    else
                        smCache.Add(key, sm);
                }
            }

            public bool ContainsKey(Guid deckGuid, int slideIndex)
            { 
                String key = deckGuid.ToString() + "-" + slideIndex.ToString();
                lock (smCache)
                {
                    if (smCache.ContainsKey(key))
                        return true;
                    else
                        return false;
                }
            }


        }

        #endregion SlideMessageCacheClass

        #region ScrollPosCache Class

        /// <summary>
		/// Cache current scroll parameters for each slide and deck.
		/// </summary>
		private class ScrollPosCache
		{
			Hashtable cache; //Key is a string made of slide deck guid and slide index, value is a ScrollParams instance

			public ScrollPosCache()
			{
				cache = new Hashtable();
			}

			#region Public Methods

			/// <summary>
			/// Add or update scroll position for the specified slide.
			/// </summary>
			/// <param name="deckGuid"></param>
			/// <param name="index"></param>
			/// <param name="pos"></param>
			/// <param name="extent"></param>
			public void Add(Guid deckGuid, Int32 index, Double pos, Double extent)
			{
				String key = deckGuid.ToString() + "-" + index.ToString();

				lock (this)
				{
					if (cache.ContainsKey(key))
					{
						cache[key]=new ScrollParams(pos,extent);
					}
					else
					{
						cache.Add(key,new ScrollParams(pos,extent));
					}
				}
			}

			/// <summary>
			/// Return scroll parameters for this slide
			/// </summary>
			/// <param name="deckGuid"></param>
			/// <param name="index"></param>
			/// <returns></returns>
			public ScrollParams Get(Guid deckGuid, Int32 index)
			{
				String key = deckGuid.ToString() + "-" + index.ToString();
				lock (this)
				{
					if (cache.ContainsKey(key))
					{
						return (ScrollParams)cache[key];
					}
					else
					{
						return new ScrollParams();
					}
				}
			}

			#endregion Public Methods
		}

		#endregion ScrollPosCache Class

		#region ScrollParams Class

		private class ScrollParams
		{

			public ScrollParams(Double pos, Double extent)
			{
				scrollPos = pos;
				scrollExtent = extent;
			}

			public ScrollParams()
			{
				//These are the default values used by Presenter.
				scrollPos = 0;
				scrollExtent = 1.5;
			}

			private Double scrollPos;
			public Double ScrollPos
			{
				get {return scrollPos;}
			}

			private Double scrollExtent;
			public Double ScrollExtent
			{
				get  {return scrollExtent;}
			}
		}

		#endregion ScrollParams Class

	}

	#region PresenterWireFormatType Enum

	public enum PresenterWireFormatType
	{
		CPNav,
		CPCapability,
		RTDocument,
		Other,
		Unknown,
        CP3,
        Video
	}

	#endregion PresenterWireFormatType Enum

	#region PresenterRoleType Enum

	public enum PresenterRoleType
	{
		Instructor,
		SharedDisplay,
		Student,
		Other,
		Unknown
	}	 
  
	#endregion PresenterRoleType Enum

}
