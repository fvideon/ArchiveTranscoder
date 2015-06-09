using System;
using System.Threading;
using System.Diagnostics;
using System.Collections;
using System.Text;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Look up and cache the items from the database that the UI will use to build the SegmentForm TreeView.
	/// </summary>
	/// The lookup can be very time consuming if we have a lot of new data, so if possible we want to start this
	/// process in a thread as soon as we launch.  Also support refresh so that the cache can be rebuilt when
	/// the database name or server name changes.
	/// 
	/// An alternate approach: use the treeview 'BeforeExpand' event to fill in only the branches of the tree that the user 
	/// actually views.  This might have been a little simpler, and nearly as effective.
	/// 
	/// Additional features possible:
	/// -do some analysis of start times and durations of video streams so that we can detect scenarios where a node
	/// has multiple video sources, but they have the same name in the database.  (Note 12/14/05: Patrick says this will
	/// be resolved by a change to the Archive service soon to permit storing 255 chars)
	/// -support an event to dynamically refresh the treeview even after the basics have been filled in.
	public class ConferenceDataCache
	{
		#region Members

		private String sqlServer;
		private String dbName;
		private bool done;
		private bool sqlConnectError;
		private Thread lookupThread;
		private ArrayList conferenceData;
		private bool stopNow;
		private int participantCount;
		private int currentParticipant;
		private ManualResetEvent working;

		#endregion Members

		#region Ctor/Dispose

		public ConferenceDataCache(String sqlServer, String dbName)
		{
			working = new ManualResetEvent(true);
			this.sqlServer = sqlServer;
			this.dbName = dbName;
			done = false;
			sqlConnectError = false;
			lookupThread = null;
			conferenceData = null;
			stopNow = false;
			participantCount = 0;
			currentParticipant = 0;
		}

		public void Dispose()
		{
			//kill the lookup thread, if any
			this.stopThread();
		}

		#endregion Ctor/Dispose

		#region Properties

		/// <summary>
		/// The results.  A list of ConferenceWrapper.  Don't access until the lookup is complete.
		/// </summary>
		public ArrayList ConferenceData
		{
			get {return conferenceData;}
		}

		/// <summary>
		/// True when there is no lookup thread in progress
		/// </summary>
		public bool Done
		{
			get {return done;}
		}
		
		/// <summary>
		/// True if the last lookup resulted in a Sql connect (or other SQL) exception.
		/// </summary>
		public bool SqlConnectError
		{
			get {return sqlConnectError;}
		}

		/// <summary>
		/// The system hosting the ArchiveService database
		/// </summary>
		public String SqlServer
		{
			get {return sqlServer;}
			set {sqlServer = value;}
		}

		/// <summary>
		/// The name of the ArchiveService database
		/// </summary>
		public String DbName
		{
			get {return dbName;}
			set {dbName = value;}
		}

		/// <summary>
		/// The caller can wait on this to get notification when the lookup is done.
		/// </summary>
		public ManualResetEvent Working
		{
			get {return working;}
		}

		/// <summary>
		/// Total participants in the current or most recent lookup.
		/// </summary>
		public int ParticipantCount
		{
			get {return participantCount;}
		}

		/// <summary>
		/// Number of the most recently looked up participant.
		/// To support the progress bar.
		/// </summary>
		public int CurrentParticipant
		{
			get {return currentParticipant;}
		}

		#endregion Properties

		#region Public Methods

		/// <summary>
		/// Start or restart the process of filling cache.
		/// </summary>
		public void Lookup()
		{
			//kill lookup thread if it is running
			this.stopThread();

			done=false;
			sqlConnectError=false;
			if ((sqlServer == null) || (sqlServer.Trim() == "") ||
				(dbName==null) || (dbName.Trim() == ""))
			{
				done=true;
				sqlConnectError = true;
				return;
			}

			participantCount = 0;
			currentParticipant = 0;

			//start a new lookup thread;
			lookupThread = new Thread(new ThreadStart(lookupThreadProc));
			lookupThread.Name = "ConferenceDataCache Lookup Thread";
			stopNow = false;
			lookupThread.Start();
		}

		/// <summary>
		/// Stop a lookup in progress, if any.
		/// </summary>
		public void Stop()
		{
			this.stopThread();
		}

		#endregion Public Methods

		#region Private Methods

		private void lookupThreadProc()
		{
			working.Reset();

			if (!DatabaseUtility.CheckConnectivity())
			{
				done=true;
				sqlConnectError = true;
				working.Set();
				return;
			}

			Conference[] confs = DatabaseUtility.GetConferences();
			this.conferenceData = new ArrayList();

			participantCount = DatabaseUtility.CountParticipants();
			if (participantCount<1)
			{
				done=true;
				sqlConnectError = false;
				working.Set();
				return;	
			}

			currentParticipant = 0;

			foreach (Conference c in confs)
			{
				if (stopNow)
					break;
				ConferenceWrapper cw = new ConferenceWrapper();
				cw.Conference = c;
				cw.Description = "Conference: " + c.Start.ToString() + " - " + c.Description;
				cw.Participants = new ArrayList();

				Participant[] participants = DatabaseUtility.GetParticipants( c.ConferenceID );
				foreach( Participant part in participants )
				{
					if (stopNow)
						break;
					ParticipantWrapper pw = new ParticipantWrapper();

                    pw.Description = "Participant: " + part.CName;
					pw.Participant = part;
					pw.StreamGrouper = new StreamGrouper(part.CName);

					//Here is the slow part:
					Stream[] streams = DatabaseUtility.GetStreamsFaster(part.ParticipantID);
					currentParticipant++;

					if (streams==null)
						continue;

					foreach (Stream str in streams)
					{
						pw.StreamGrouper.AddStream(str);
					}

					//Only show the participants that sent some supported data
					if (pw.StreamGrouper.GroupCount > 0)
						cw.Participants.Add(pw);
				}

				conferenceData.Add(cw);
			}
			done=true;
			working.Set();
		}
		
		private void stopThread()
		{
			stopNow = true;
			if (lookupThread != null)
			{
				if (lookupThread.IsAlive)
				{
					if (!lookupThread.Join(2000))
					{
						Debug.WriteLine("Aborting ConferenceDataCache lookup thread.");
						lookupThread.Abort();
					}
				}
				lookupThread = null;
			}
			done=true;
			working.Set();
		}

		#endregion Private Methods
	}

	#region ConferenceWrapper Class

	/// <summary>
	/// Simple class to help mapping between database and SegmentForm TreeView
	/// </summary>
	public class ConferenceWrapper
	{
		public Conference Conference;
		public String Description;
		public ArrayList Participants;
	}

	#endregion ConferenceWrapper Class

	#region ParticipantWrapper Class

	/// <summary>
	/// Simple class to help mapping between database and SegmentForm TreeView
	/// </summary>
	public class ParticipantWrapper
	{
		public Participant Participant;
		public String Description;
		public StreamGrouper StreamGrouper;
	}

	#endregion ParticipantWrapper Class

	#region StreamGrouper Class

	/// <summary>
	/// Take inputs about a set of streams for one participant, and group them into the collection
	/// appropriate to display in the treeview.  
	/// </summary>
	/// -All audio streams belong in one line
	/// -Each video camera or WMV source should appear on a different video line 
	/// -Each presentation wire format (and role?) should appear on a separate presentation line
	public class StreamGrouper
	{
		private String cname;
		//private StreamGroup audioGroup;
		private ArrayList audioGroups;
		private ArrayList videoGroups;
		private ArrayList presGroups;

		public StreamGrouper(String cname)
		{
			this.cname = cname;
			audioGroups  = new ArrayList();
			videoGroups = new ArrayList();
			presGroups = new ArrayList();
		}

		public int GroupCount
		{
			get 
			{
				int count = videoGroups.Count + audioGroups.Count + presGroups.Count;
				return count;
			}
		}

		/// <summary>
		/// Return an arraylist of StreamGroup in the order we want to display them.
		/// </summary>
		/// <returns></returns>
		public ArrayList GetStreamList()
		{
			ArrayList streamList = new ArrayList();
			foreach (StreamGroup sg in audioGroups)
			{
				streamList.Add(sg);
			}
			foreach (StreamGroup sg in videoGroups)
			{
				streamList.Add(sg);
			}
			foreach (StreamGroup sg in presGroups)
			{
				streamList.Add(sg);
			}
			return streamList;
		}

		/// <summary>
		/// Add this stream to the appropriate group, or create a new group as necessary.
		/// </summary>
		/// <param name="stream"></param>
		public void AddStream(Stream stream)
		{
			if (stream.Payload == "dynamicAudio")
			{
				/// if stream name matches an existing group, just add it
				/// otherwise make a new group.
				bool matchfound = false;
				foreach (StreamGroup sg in audioGroups)
				{
					if (sg.IsInAudioGroup(stream))
					{
						sg.AddStream(stream);
						matchfound = true;
						break;
					}
				}
				if (!matchfound)
				{
					StreamGroup newsg = new StreamGroup(this.cname,"",stream.Payload);
					newsg.AddStream(stream);
					audioGroups.Add(newsg);
				}
			}
			else if (stream.Payload == "dynamicVideo")
			{
				/// if stream name matches an existing group, just add it
				/// otherwise make a new group.
				bool matchfound = false;
				foreach (StreamGroup sg in videoGroups)
				{
					if (sg.IsInVideoGroup(stream))
					{
						sg.AddStream(stream);
						matchfound = true;
						break;
					}
				}
				if (!matchfound)
				{
					StreamGroup newsg = new StreamGroup(this.cname,"",stream.Payload);
					newsg.AddStream(stream);
					videoGroups.Add(newsg);
				}
			}
			else if ((stream.Payload == "dynamicPresentation") ||  //CP running standalone uses dynamicPresentation
					(stream.Payload == "RTDocument"))			   //CP as a capability or CXP Presenter uses RTDocument
			{
				PresenterWireFormatType wireFormat;
				PresenterRoleType role;
                string presenterName;
				if (PresentationMgr.GetPresentationType(stream, out wireFormat, out role, out presenterName))
				{
                    ///PRI2: We need a better solution here:  We would like to filter out all non-instructors, but
                    ///we can't afford to look at every frame.
					if ((wireFormat == PresenterWireFormatType.CPNav) || 
						(wireFormat == PresenterWireFormatType.CPCapability) ||
						(wireFormat == PresenterWireFormatType.RTDocument) ||
                        (wireFormat == PresenterWireFormatType.CP3))
						// && (role == PresenterRoleType.Instructor)) 
					{
						//add to existing group if there is a match
						bool matchfound = false;
						foreach (StreamGroup sg in presGroups)
						{
							if (sg.IsInPresGroup(stream,wireFormat,role))
							{
								sg.AddStream(stream);
								matchfound = true;
								break;
							}
						}
						//if no match, create a new group
						if (!matchfound)
						{
							StreamGroup newsg = new StreamGroup(this.cname,"",stream.Payload);
							newsg.AddStream(stream,wireFormat,role);
							presGroups.Add(newsg);
						}
					}
				}
			}
		}
	}

	#endregion StreamGrouper Class

	#region StreamGroup Class

	/// <summary>
	/// Represents one or more streams that will be described by a single line in the treeview
	/// </summary>
	public class StreamGroup
	{
		private ArrayList streams;
		private String payload;
		private String cname;
		private PresenterWireFormatType wireFormat = PresenterWireFormatType.Unknown;
		private String name;
		private PresenterRoleType role = PresenterRoleType.Unknown;
		
		public String Payload
		{
			get {return payload;}
		}

		public PresenterWireFormatType Format
		{
			get {return wireFormat;}
		}

		public String Cname
		{
			get {return cname;}
		}

		public String Name
		{
			get {return name;}
		}

		public StreamGroup(String cname, String name, String payload)
		{
			streams = new ArrayList();
			this.cname = cname;
			this.payload = payload;
			role = PresenterRoleType.Other;
			wireFormat = PresenterWireFormatType.Other;
		}

		/// <summary>
		/// Ctors for creating a streamgroup from an existing batch file
		/// </summary>
		/// <param name="videoDescriptor"></param>
		public StreamGroup(ArchiveTranscoderJobSegmentVideoDescriptor videoDescriptor)
		{
			streams = new ArrayList();
			this.cname = videoDescriptor.VideoCname;
			this.name = videoDescriptor.VideoName;
			this.payload = "dynamicVideo";
		}

		public StreamGroup(ArchiveTranscoderJobSegmentAudioDescriptor audioDescriptor)
		{
			streams = new ArrayList();
			this.cname = audioDescriptor.AudioCname;
			this.name = audioDescriptor.AudioName;
			this.payload = "dynamicAudio";
		}

		public StreamGroup(ArchiveTranscoderJobSegmentPresentationDescriptor presDescriptor)
		{
			streams = new ArrayList();
			this.payload = "dynamicPresentation";
			this.cname = presDescriptor.PresentationCname;

			this.wireFormat = PresenterWireFormatType.Unknown;
			foreach (PresenterWireFormatType t in Enum.GetValues(typeof(PresenterWireFormatType)))
			{
				if (presDescriptor.PresentationFormat==t.ToString())
				{
					this.wireFormat = t;
					payload = Utility.formatToPayload(t).ToString();
					break;
				}
			}

		}

		public ArchiveTranscoderJobSegmentPresentationDescriptor ToPresentationDescriptor()
		{
			ArchiveTranscoderJobSegmentPresentationDescriptor d = new ArchiveTranscoderJobSegmentPresentationDescriptor();
            d.PresentationCname = this.cname;
            if (this.payload == "dynamicPresentation") {
                d.PresentationFormat = this.wireFormat.ToString();
            }
            else if (this.payload == "RTDocument") {
                d.PresentationFormat = this.wireFormat.ToString();
            }
            else if (this.payload == "dynamicVideo") {
                d.PresentationFormat = "Video";
                d.VideoDescriptor = this.ToVideoDescriptor();
            }
			return d;
		}

		public ArchiveTranscoderJobSegmentVideoDescriptor ToVideoDescriptor()
		{
			ArchiveTranscoderJobSegmentVideoDescriptor d = new ArchiveTranscoderJobSegmentVideoDescriptor();
			d.VideoCname = this.cname;
			d.VideoName = this.name;
			return d;
		}

		public ArchiveTranscoderJobSegmentAudioDescriptor ToAudioDescriptor()
		{
			ArchiveTranscoderJobSegmentAudioDescriptor d = new ArchiveTranscoderJobSegmentAudioDescriptor();
			d.AudioCname = this.cname;
			d.AudioName = this.name;
			return d;
		}

		public void AddStream(Stream stream)
		{
			streams.Add(stream);
			this.name = stream.Name;
		}

		public void AddStream(Stream stream, PresenterWireFormatType wireFormat, PresenterRoleType role)
		{
			this.wireFormat = wireFormat;
			this.role = role;
			this.AddStream(stream);
		}

		/// <summary>
		/// Lacking a better way, we group video streams if the names match.
		/// This is not very robust since names are truncated to 40 chars in the database, and the unique
		/// part is a number in square brackets at the end which may be tuncated.  
		/// </summary>
		/// <param name="s"></param>
		/// <returns></returns>
		public bool IsInVideoGroup(Stream s)
		{
			if ((s.Payload==this.payload) && (this.payload=="dynamicVideo"))
			{
				foreach (Stream str in streams)
				{
					if (str.Name==s.Name)
					{
						return true;
					}
				}
			}
			return false;
		}

		/// <summary>
		/// Return true if stream is audio and has a name that matches this group.
		/// </summary>
		/// <param name="s"></param>
		/// <returns></returns>
		public bool IsInAudioGroup(Stream s)
		{
			if ((s.Payload==this.payload) && (this.payload=="dynamicAudio"))
			{
				foreach (Stream str in streams)
				{
					if (str.Name==s.Name)
					{
						return true;
					}
				}
			}
			return false;
		}


		/// <summary>
		/// Return true if this stream matches the existing format/role in the group.
		/// </summary>
		/// <param name="s"></param>
		/// <param name="wireFormat"></param>
		/// <param name="role"></param>
		/// <returns></returns>
		public bool IsInPresGroup(Stream s, PresenterWireFormatType wireFormat, PresenterRoleType role)
		{
			if ((s.Payload==this.payload) && ((this.payload=="dynamicPresentation") || (this.payload=="RTDocument")))
			{
				if ((wireFormat == this.wireFormat) &&
					(role == this.role))
				{
					return true;
				}
			}
			return false;
		}

		/// <summary>
		/// Return the string representation of the StreamGroup that we want to display in the ListBox
		/// </summary>
		/// <returns></returns>
		public override string ToString()
		{
			if ((this.payload=="dynamicPresentation") || (this.payload=="RTDocument"))
			{
				if (this.wireFormat == PresenterWireFormatType.CPCapability)
				{
					return this.cname + " - Classroom Presenter Capability";				
				}
				else if (this.wireFormat == PresenterWireFormatType.CPNav)
				{
					return this.cname + " - Classroom Presenter";				
				}
				else if (this.wireFormat == PresenterWireFormatType.RTDocument)
				{
					return this.cname + " - RTDocument (CXP Presentation)";
				}
                else if (this.wireFormat == PresenterWireFormatType.CP3) {
                    return "CP3 {" + this.cname + "}";
                }
            }
			else if ((this.name != null) && (this.name.Trim() != ""))
			{
				return this.cname + " - " + this.name;
			}
			return this.cname;
		}

		/// <summary>
		/// A longer description for the treeview.
		/// </summary>
		/// <returns></returns>
		public String ToTreeViewString()
		{
			StringBuilder sb = new StringBuilder();
			if (this.payload=="dynamicAudio")
			{
				sb.Append("Audio");
			}
			else if (this.payload=="dynamicVideo")
			{
				sb.Append("Video");
			}
			else if (this.payload=="dynamicPresentation")
			{
				sb.Append("Presentation");
			}
			else if (this.payload=="RTDocument")
			{
				sb.Append("Presentation");
			}
			else
			{
				return "";
			}
			
			sb.Append(" (");
			if (PresenterRoleType.Instructor == role)
			{
				sb.Append("Instr; ");
			}
			if (this.streams.Count > 1)
			{
				sb.Append(this.streams.Count.ToString() + " streams; ");
			}

			long duration = 0;
			foreach (Stream s in streams)
			{
				duration += s.Seconds;
			}

			sb.Append("duration=" + TimeSpan.FromSeconds(duration).ToString());
			sb.Append(")");

			if (streams.Count >0)
			{
				if (this.payload=="dynamicAudio")
				{
					sb.Append(" " + ((Stream)streams[0]).Name);
				}
				else if (this.payload=="dynamicVideo")
				{
					sb.Append(" " + ((Stream)streams[0]).Name);
				}
				else
				{
					sb.Append(" " + ((Stream)streams[0]).Name);
				}
			}

			return sb.ToString();
		}
	}

	#endregion StreamGroup Class

}
