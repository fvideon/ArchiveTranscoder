using System;

namespace ArchiveTranscoder
{
	#region ProfileWrapper Class

	/// <summary>
	/// Simple class to wrap WM Profiles for easy use in a combobox
	/// </summary>
	public class ProfileWrapper
	{
		public enum EntryType
		{ System, Custom, NoRecompression }

		private EntryType entryType;
		private String displayText;
		private String profileValue;

		public ProfileWrapper(EntryType entryType, String displayText, String profileValue)
		{
			this.entryType = entryType;
			this.displayText = displayText;
			this.profileValue = profileValue;
		}

		public String ProfileValue
		{
			get {return profileValue;}
		}

		public override string ToString()
		{
			return displayText;
		}

	}

	#endregion ProfileWrapper Class

	#region SegmentWrapper Class

	/// <summary>
	/// Wrap ArchiveTranscoderJobSegment for use as a ListBox Item.
	/// </summary>
	public class SegmentWrapper:IComparable
	{
		private ArchiveTranscoderJobSegment segment;
		private DateTime starttime;
		public ArchiveTranscoderJobSegment Segment
		{
			get {return segment;}
		}

		public SegmentWrapper(ArchiveTranscoderJobSegment segment)
		{
			this.segment = segment;
			starttime = DateTime.Parse(segment.StartTime);
		}

		public override string ToString()
		{
			int cvideo,caudio,cpres;
			cvideo=caudio=cpres=0;
			if ((segment.VideoDescriptor != null) &&
				(segment.VideoDescriptor.VideoCname != null) && 
				(segment.VideoDescriptor.VideoCname != ""))
				cvideo++;
			if ((segment.AudioDescriptor != null) && (segment.AudioDescriptor.Length > 0))
			{
				caudio = segment.AudioDescriptor.Length;
			}
			if ((segment.PresentationDescriptor != null) &&
				(segment.PresentationDescriptor.PresentationCname != null) && 
				(segment.PresentationDescriptor.PresentationCname != ""))
				cpres++;

            String videoSubstring = cvideo.ToString() + " video,";
            String presentationSubstring = cpres.ToString() + " presentation.";
            String audioSubstring = caudio.ToString() + " audio,";
            if (Utility.SegmentFlagIsSet(segment, SegmentFlags.SlidesReplaceVideo))
            {
                videoSubstring = "video from presentation,";
                audioSubstring = caudio.ToString() + " audio.";
                presentationSubstring = "";
            }

			DateTime start = DateTime.Parse(segment.StartTime);
			TimeSpan duration = DateTime.Parse(segment.EndTime) - start;
			String durationString = duration.ToString();
			if (durationString.LastIndexOf(".") > 0)
				durationString = durationString.Substring(0,durationString.LastIndexOf("."));

			return "Start: " + Utility.GetLocalizedDateTimeString(start, Constants.shorttimeformat) +
                "; Duration: " + durationString + "; " + videoSubstring + audioSubstring + presentationSubstring;
		}

		public string Description()
		{
			String ret = "Start: " + segment.StartTime.ToString() + "\n\r" +
				"End: " + segment.EndTime.ToString() + "\n\r";
			if ((segment.VideoDescriptor != null) &&
				(segment.VideoDescriptor.VideoCname != null) && 
				(segment.VideoDescriptor.VideoCname != ""))
				ret += "Video: " + segment.VideoDescriptor.VideoCname + "\n\r";
			if ((segment.AudioDescriptor != null) && (segment.AudioDescriptor.Length > 0))
			{
				for (int i=0; i< segment.AudioDescriptor.Length; i++)
				{
					ret += "Audio: " + segment.AudioDescriptor[i].AudioCname + "\n\r";
				}
			}
			if ((segment.PresentationDescriptor != null) &&
				(segment.PresentationDescriptor.PresentationCname != null) && 
				(segment.PresentationDescriptor.PresentationCname != ""))
				ret += "Presentation: " + segment.PresentationDescriptor.PresentationCname;

			return ret;
		}
		#region IComparable Members

		public int CompareTo(object obj)
		{
			if(obj is SegmentWrapper) 
			{
				SegmentWrapper sw = (SegmentWrapper) obj;
				return starttime.CompareTo(sw.starttime);
			}
			throw new ArgumentException("object is not a SegmentWrapper");
		}

		#endregion
	}
	#endregion SegmentWrapper Class
}
