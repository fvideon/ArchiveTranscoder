using System;
using System.Collections;
using System.Text;
using MSR.LST.MDShow;

namespace ArchiveTranscoder
{
	#region AudioCompatibilityMgr Class
	/// <summary>
	/// Facilitate audio media type checking and logging, and the exclusion of data 
	/// with incompatible types.  Implement a simple voting mechanism to determine
	/// which media type is dominant among all types found in this segment.
	/// </summary>
	public class AudioCompatibilityMgr
	{
		private ArrayList audioTypeMonitors;

		public AudioCompatibilityMgr()
		{
			audioTypeMonitors = new ArrayList();
		}

		/// <summary>
		/// Calculate the list of FileStreamPlayer guids that are incompatible with the dominant media type.
		/// </summary>
		/// <returns></returns>
		public ArrayList GetIncompatibleGuids()
		{
			ArrayList al = new ArrayList();
			ulong maxduration = 0;
			int maxdurationindex = 0;
			for(int i=0;i<audioTypeMonitors.Count;i++)
			{
				if (((AudioTypeMonitor)(audioTypeMonitors[i])).Duration > maxduration)
				{
					maxduration = ((AudioTypeMonitor)(audioTypeMonitors[i])).Duration;
					maxdurationindex = i;
				}
			}
			StringBuilder warnings = new StringBuilder();
			for(int i=0;i<audioTypeMonitors.Count;i++)
			{
				if (i!=maxdurationindex)
				{
					foreach(Guid g in ((AudioTypeMonitor)(audioTypeMonitors[i])).Streams.Keys)
					{
						al.Add(g);
					}
				}
			}
			return al;
		}

		/// <summary>
		/// Record the MediaType vote and stream details for this stream.
		/// </summary>
		/// <param name="mt"></param>
		/// <param name="duration"></param>
		/// <param name="guid"></param>
		/// <param name="cname"></param>
		/// <param name="starttime"></param>
		public void Check(MediaTypeWaveFormatEx mt, ulong duration, Guid guid, String cname, String name, long starttime,int streamID)
		{
			bool match = false;
			foreach (AudioTypeMonitor mon in audioTypeMonitors)
			{
				if (mon.Matches(mt))
				{
					mon.AddToDuration(duration,guid,cname,name,starttime,streamID);
					match = true;
					break;
				}
			}
			if (!match)
			{
				audioTypeMonitors.Add(new AudioTypeMonitor(mt,duration,guid,cname,name,starttime,streamID));
			}
		}

		/// <summary>
		/// Get a streamID from the AudioTypeMonitor with highest duration.
		/// </summary>
		/// <returns></returns>
		public int GetCompatibleStreamID()
		{
			ulong maxduration = 0;
			int maxdurationindex = 0;
			for(int i=0;i<audioTypeMonitors.Count;i++)
			{
				if (((AudioTypeMonitor)(audioTypeMonitors[i])).Duration > maxduration)
				{
					maxduration = ((AudioTypeMonitor)(audioTypeMonitors[i])).Duration;
					maxdurationindex = i;
				}
			}
			return 	((AudioTypeMonitor)(audioTypeMonitors[maxdurationindex])).StreamID;
		}


		/// <summary>
		/// Find the type with the highest vote count.
		/// </summary>
		/// <returns></returns>
		public MediaTypeWaveFormatEx GetDominantType()
		{
			ulong maxduration = 0;
			MediaTypeWaveFormatEx dominantType = null;
			foreach (AudioTypeMonitor mon in audioTypeMonitors)
			{
				if (mon.Duration > maxduration)
				{
					maxduration = mon.Duration;
					dominantType = mon.MT;
				}
			}
			return dominantType;
		}

		/// <summary>
		/// Return details about incompatible streams.
		/// </summary>
		/// <returns></returns>
		public String GetWarningString()
		{
			ulong maxduration = 0;
			int maxdurationindex = 0;
			//Get the index of the most prevalent media type:
			for(int i=0;i<audioTypeMonitors.Count;i++)
			{
				if (((AudioTypeMonitor)(audioTypeMonitors[i])).Duration > maxduration)
				{
					maxduration = ((AudioTypeMonitor)(audioTypeMonitors[i])).Duration;
					maxdurationindex = i;
				}
			}
			//Build a list of sources of the other incompatible media types:
			StringBuilder warnings = new StringBuilder();
			for(int i=0;i<audioTypeMonitors.Count;i++)
			{
				if (i!=maxdurationindex)
				{
					warnings.Append(((AudioTypeMonitor)(audioTypeMonitors[i])).ToString());
				}
			}

			if (warnings.Length>0)
				return "The following audio streams were found to have incompatible media types, and will not be included in the mix: " +
					warnings.ToString();
			else
				return "";
		}
	}

	#endregion AudioCompatibilityMgr Class

	#region AudioTypeMonitor Class

	/// <summary>
	/// Encapsulate a Audio MediaType and a duration to facilitate finding out which is
	/// the most used uncompressed audio media type in a segment. 
	/// </summary>
	public class AudioTypeMonitor
	{
		private Hashtable streams;
		private ulong duration;
		private MediaTypeWaveFormatEx mt;
		private int streamID;

		/// <summary>
		/// This Monitor's MediaType
		/// </summary>
		public MediaTypeWaveFormatEx MT
		{
			get { return mt; }
		}

		/// <summary>
		/// Hashtable of streams that are compatible with this media type.
		/// Guid is key, StreamData is value.
		/// </summary>
		public Hashtable Streams
		{
			get { return streams; }
		}

		/// <summary>
		/// This sum of durations of all streams successfully added to this Monitor.
		/// </summary>
		public ulong Duration
		{
			get {return duration;}
		}

		/// <summary>
		/// A sample streamID belonging to this Monitor
		/// </summary>
		public int StreamID
		{
			get {return streamID;}
		}

		public AudioTypeMonitor(MediaTypeWaveFormatEx mt, ulong duration, Guid guid, String cname, String name, long starttime, int streamID)
		{
			this.mt = mt;
			this.duration = duration;
			this.streamID = streamID;
			streams = new Hashtable();
			streams.Add(guid,new StreamData(guid,cname,name,starttime,streamID));
		}

		/// <summary>
		/// Return a list of the streams in this Monitor
		/// </summary>
		/// <returns></returns>
		public override string ToString()
		{
			StringBuilder s = new StringBuilder();
			foreach (StreamData sd in streams.Values)
			{
				s.Append(sd.ToString() + "\r\n");
			}
			return s.ToString();
		}

		/// <summary>
		/// Return true if the type is compatible for PCM mixing.  Also return
		/// true if the type is null.
		/// </summary>
		/// <param name="mt2"></param>
		/// <returns></returns>
		public bool Matches(MediaTypeWaveFormatEx mt2)
		{
			if ((mt == null) || (mt2 == null))
				return true;  //one of them is unassigned.
			
			if ((mt.MajorType == mt2.MajorType) &&
				(mt.SubType == mt2.SubType) &&
                (mt.WaveFormatEx.SamplesPerSec == mt2.WaveFormatEx.SamplesPerSec) &&
				(mt.WaveFormatEx.BitsPerSample == mt2.WaveFormatEx.BitsPerSample))
				return true;

			return false;
		}

		/// <summary>
		/// Add this stream's duration to the cumulative total for this media type, and record stream details for logging.
		/// </summary>
		/// <param name="addTime"></param>
		/// <param name="guid"></param>
		/// <param name="cname"></param>
		/// <param name="starttime"></param>
		public void AddToDuration(ulong addTime, Guid guid, String cname, String name, long starttime, int streamID)
		{
			duration += addTime;
			if (!streams.ContainsKey(guid))
			{
				streams.Add(guid, new StreamData(guid,cname,name,starttime,streamID));
			}
		}

		#region StreamData Class
		/// <summary>
		/// Simple class to maintain a few properties about a stream.  Mainly used to facilitate logging.
		/// </summary>
		private class StreamData
		{
			private Guid guid;
			private String cname;
			private String name;
			private DateTime starttime;
			private int streamID;

			public int StreamID
			{
				get {return streamID; }
			}

			public StreamData(Guid guid, String cname, String name, long starttime, int streamID)
			{
				this.guid = guid;
				this.cname = cname;
				this.name = name;
				this.starttime = new DateTime(starttime);
				this.streamID = streamID;
			}

			public override string ToString()
			{
				if (Utility.isSet(name))
					return "cname=" + cname + " - " + name + " start time=" + starttime.ToString();
				else
					return "cname=" + cname + " start time=" + starttime.ToString();
			}
		}
		#endregion StreamData Class
	}

	#endregion AudioTypeMonitor Class
}
