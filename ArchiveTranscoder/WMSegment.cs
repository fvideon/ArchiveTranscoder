using System;
using System.Diagnostics;
using System.Collections;
using MSR.LST;
using MSR.LST.Net.Rtp;
using MSR.LST.MDShow;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Encode one segment of a job.
	/// </summary>
	/// 
	/// By definition a segment consists of one time range, one video cname, an optional video
	/// camera name, one or more audio cnames, and an optional presentation source. For each cname
	/// there may be one or more non-overlapping streams.  Overlapping audio streams across all audio cnames
	/// are mixed here.  
	/// 
	/// The main method call supported for WMSegment is Encode.  Encoding can be performed with or without
	/// recompression.   In some circumstances recompression can be applied to audio but not to video.
	/// 
	/// If encoding with recompression, we need to receive uncompressed samples from Audio/Video Processors.
	/// Without recompression we cannot mix audio, so are limited to only one audio source.  We also
	/// can't tolerate changes in the media types without recompression.  
	/// 
	/// After preparing Audio/video processors and determining recompression, iterate over the 
	/// timespan writing Samples or StreamSamples to WMWriter as appropriate.
	/// 
	public class WMSegment
	{
		#region Members

        private ArchiveTranscoderJob job;
		private ArchiveTranscoderJobSegment segment;	//The part of the batch defining params for this segment.
		private WMWriter writer;						//Windows Media Writer object
		private DateTime startTime;						//Segment start time parsed as a DateTime
		private DateTime endTime;						//Segment end time parsed as a DateTime
		private StreamMgr videoStream;					//Reference to object that handles video streams
        private SlideStreamMgr slideStream;             //Generate video derived from a presentation stream
		private StreamMgr[] audioStreams;				//References to objects that handle audio streams
		private PresentationMgr presentationMgr;		//Handles optional presentation data
		private AudioMixer mixer;						//Audio mixer class
		private ProgressTracker progressTracker;		//Passes status updates to the UI.
		private bool cancel;							//Flags cancellation of an encode job in progress
        private bool useSlideStream;                     //Indicates that we want to use slide images in place of the video
		private long offset;							//Relative time of the start of this segment, in ticks from jobStart.
		private long jobStart;							//Absolute time of the start of the first segment in the job in ticks.
		private long actualSegmentStart;				//The absolute time in ticks of the first WM Write operation for audio or video in this segment.
		private long actualSegmentEnd;					//The absolute time in ticks of the last WM Write operation for audio or video in this segment.
		private LogMgr logMgr;							//Writes log messages
        private ProfileData profileData;                //In the norecompression case, this contains the compressed media types
        private bool norecompression;                   //Indicates that the segment should write compressed samples
        private bool audiorecompression;
        private WMSegment m_PreviousSegment;

		#endregion Members

		#region Properties


		/// <summary>
		/// The timespan in ticks from the first actual AV sample written to the last.  
		/// This is only valid after the Encode method has
		/// completed.
		/// </summary>
		public long ActualWriteDuration
		{
			get {return this.actualSegmentEnd-this.actualSegmentStart;}
		}

		/// <summary>
		/// The absolute time in ticks of the first AV sample written in this segment.  
		/// This is only valid after the Encode method has
		/// completed.
		/// </summary>
		public long ActualSegmentStart
		{
			get {return this.actualSegmentStart;}
		}

        /// <summary>
        /// absolute time in ticks of the last AV sample written in this segment.
        /// Only valid after the encode method has completed
        /// </summary>
        public long ActualSegmentEnd {
            get { return this.actualSegmentEnd; }
        }

		/// <summary>
		/// If there are RTDocuments in the presentation data in the segment, this property returns
		/// them, keyed by Identifier (Guid).  This is only valid after the Encode method has
		/// completed.
		/// </summary>
		public Hashtable RTDocuments
		{
			get
			{				
				if (presentationMgr != null)
				{
					return presentationMgr.RTDocuments;
				}
				else
					return null;
			}
		}

        public Hashtable SlideMessages
        {
            get
            {
                if (presentationMgr != null)
                {
                    return presentationMgr.SlideMessages;
                }
                else
                    return null;
            }
        }

		/// <summary>
		/// If there is presentation data in the segment, this property returns
		/// a list of deck titles found.  This is only valid after the Encode method
		/// has completed.
		/// </summary>
		public Hashtable DeckTitles
		{
			get
			{
				if (presentationMgr != null)
				{
					return presentationMgr.DeckTitles;
				}
				else
					return null;
			}
		}

		/// <summary>
		/// True if supported presentation data was found for this segment.
		/// This is only valid after the Encode method has completed.
		/// </summary>
		public bool ContainsPresentationData
		{
			get 
			{
				if (presentationMgr != null)
				{
					if (presentationMgr.ContainsData)
					{
						return true;
					}
				}
				return false;
			}
		}

        public PresentationFromVideoMgr FrameGrabMgr {
            get {
                if (this.presentationMgr != null) {
                    return this.presentationMgr.FrameGrabMgr;
                }
                return null;
            }
        }

        public string FrameGrabDirectory { 
            get {
                if ((this.presentationMgr != null) && (this.presentationMgr.FrameGrabMgr != null)) {
                    return this.presentationMgr.FrameGrabMgr.TempDir;
                }
                return null;
            }
        }

        public Guid FrameGrabGuid {
            get {
                if ((this.presentationMgr != null) && (this.presentationMgr.FrameGrabMgr != null)) {
                    return this.presentationMgr.FrameGrabMgr.DeckGuid;
                }
                return Guid.Empty;
            }
        }

        public List<double> FrameGrabDifferenceMetrics {
            get {
                if ((this.presentationMgr != null) && (this.presentationMgr.FrameGrabMgr != null)){
                    return this.presentationMgr.FrameGrabMgr.FrameGrabDifferenceMetrics;
                }
                return null;
            }
        }


		#endregion Properties

		#region Ctor

		/// <summary>
		/// To use: Prepare Windows Media Writer and other objects, construct, call the Encode method, 
		/// use properties and other methods to get results, release, close Windows Media Writer.
		/// </summary>
		/// <param name="segment">Batch params for this segment.  We assume all fields have already been validated.</param>
		/// <param name="jobStart">Absolute start time of the first segment of the job in ticks</param>
		/// <param name="offset">The total timespan in ticks of all previous segments in this job</param>
		/// <param name="wmWriter">Windows Media Writer object to write to.  We assume it has been preconfigured and started.</param>
		/// <param name="progressTracker">Where to put UI status updates</param>
		/// <param name="logMgr">Where to put log messages</param>
        /// <param name="profileData">compressed media types we are writing to in the norecompression case</param>
		public WMSegment(ArchiveTranscoderJob job, ArchiveTranscoderJobSegment segment, 
            long jobStart, long offset, WMWriter wmWriter, 
			ProgressTracker progressTracker, LogMgr logMgr, ProfileData profileData, bool norecompression,
            WMSegment previousSegment)
		{
			cancel = false;
			this.logMgr = logMgr;
            this.job = job;
			this.segment = segment;
			this.offset = offset;
			this.jobStart = jobStart;
			this.progressTracker = progressTracker;
            this.profileData = profileData;
            this.norecompression = norecompression;
            this.m_PreviousSegment = previousSegment;
			videoStream = null;
            slideStream = null;
			audioStreams = null;
			presentationMgr = null;
			writer = wmWriter;
			startTime = endTime = DateTime.MinValue;
			startTime = DateTime.Parse(segment.StartTime);
			endTime = DateTime.Parse(segment.EndTime);
            useSlideStream = false;
            if (Utility.SegmentFlagIsSet(segment, SegmentFlags.SlidesReplaceVideo))
            {
                useSlideStream = true;
            }
		}

		#endregion Ctor

		#region Public Methods

		/// <summary>
		/// Write audio, video and presentation data for this segment.  This can be a long-running process.
		/// It can be cancelled with the Stop method.
		/// </summary>
		/// <param name="noRecompression"></param>
		/// <returns>A message string indicates a serious problem.  Null for normal termination.</returns>
		public String Encode()
		{
			if (cancel)
				return null;

			if ((startTime == DateTime.MinValue) || (endTime == DateTime.MinValue))
				return "Invalid timespan.";

            if (useSlideStream && norecompression)
            {
                return "A slide stream cannot be processed in 'no recompression' mode.";
            }

            progressTracker.EndValue = (int)((TimeSpan)(endTime - startTime)).TotalSeconds;

            if (useSlideStream)
            {
                slideStream = new SlideStreamMgr(job, segment, logMgr, 29.97, writer.FrameWidth, writer.FrameHeight); //Using slides in place of the video
            }
            else
            {
                videoStream = new StreamMgr(segment.VideoDescriptor.VideoCname, segment.VideoDescriptor.VideoName, startTime, endTime, norecompression, PayloadType.dynamicVideo);
            }

			audiorecompression = !norecompression;
			//In this case we actually do need to recompress just the audio:
			if ((norecompression) && (segment.AudioDescriptor.Length != 1))
				audiorecompression = true;

			audioStreams = new StreamMgr[segment.AudioDescriptor.Length];
			for (int i=0; i<segment.AudioDescriptor.Length; i++)
			{
				audioStreams[i] = new StreamMgr(segment.AudioDescriptor[i].AudioCname, segment.AudioDescriptor[i].AudioName,startTime,endTime,!audiorecompression,PayloadType.dynamicAudio);
			}

			if (cancel)
				return null;

			actualSegmentStart = 0;
			actualSegmentEnd = 0;

			if (norecompression)
			{
                if (useSlideStream)
                { 
                    //Not supported
                }
                else
                {
				    // Warn and log an error if a problem is detected with the media type, but try to proceed anyway.
                    videoStream.ValidateCompressedMT(profileData.VideoMediaType, logMgr);
                    // Make the last MT available to the caller to pass to the next segment to facilitate the checking.
                    //this.compressedVideoMediaType = videoStream.GetFinalCompressedVideoMediaType();
                }

				if (audioStreams.Length == 1)
				{
					//There is truly no recompression in this case.
					///as above, do the same check with the Audio MediaType.  Log a warning if the MT changed, but attempt to proceed.
					audioStreams[0].ValidateCompressedMT(profileData.AudioMediaType,logMgr);
					//this.compressedAudioMediaType = audioStreams[0].GetFinalCompressedAudioMediaType();
					progressTracker.AVStatusMessage = "Writing Raw AV";
				}
				else
				{
					//In this case we have to recompress audio in order to do the mixing, but that should be relatively quick.
					//Note that the WMSDK docs say that SetInputProps must be set before BeginWriting.  This implies that 
					//alternating between writing compressed and uncompressed samples is not supported.  Therefore we will
					//first recompress the mixed audio, then deliver compressed samples to the writer.

					for (int i=0;i<audioStreams.Length;i++)
					{
						progressTracker.AVStatusMessage = "Reading Audio (" + (i+1).ToString() + " of " + audioStreams.Length.ToString() + ")";
						if (cancel)
							return null;
						if (!audioStreams[i].ToRawWMFile(progressTracker))
						{
							return "Failed to configure a raw audio profile.";												
						}
					}

					progressTracker.AVStatusMessage = "Mixing Audio";
					mixer = new AudioMixer(audioStreams, this.logMgr);

					/// PRI3: We could tell the mixer to recompress with the previous segment's MT (if any).  
					/// For now we just use the 
					/// mixer's voting mechanism to choose the 'dominant' input (uncompressed) format, 
					/// and make the profile from one of the streams that uses that format.

					mixer.Recompress(progressTracker);
					progressTracker.AVStatusMessage = "Writing Raw AV";
				}
			}
			else // Recompress both audio and video
			{			
				//In order to recompress, we first need to convert each stream to a raw wmv/wma
                //A slide stream starts life uncompressed, so just initialize decks.
                if (!useSlideStream)
                {
                    progressTracker.AVStatusMessage = "Reading Video";
                    if (!videoStream.ToRawWMFile(progressTracker))
                    {
                        return "Failed to configure the raw video profile.";
                    }
                }
                else
                {
                    if (!slideStream.Init(progressTracker))
                    {
                        return "Failed to prepare slide decks.";
                    }
                }
				for (int i=0;i<audioStreams.Length;i++)
				{
					progressTracker.AVStatusMessage = "Reading Audio (" + (i+1).ToString() + " of " + audioStreams.Length.ToString() + ")";
					if (cancel)
						return null;
					if (!audioStreams[i].ToRawWMFile(progressTracker))
					{
						return "Failed to configure a raw audio profile.";					
					}
				}

				mixer = new AudioMixer(audioStreams, this.logMgr);

				writer.GetInputProps();
				//The SDK allows us to reconfigure the MediaTypes on the fly if we are writing uncompressed samples.
				//We do this at the beginning of every segment, even though most of the time it is probably the same MT.
				writer.ConfigAudio(mixer.UncompressedAudioMediaType);
                if (useSlideStream)
                {
                    writer.ConfigVideo(slideStream.UncompressedMediaType);
                    //writer.DebugPrintVideoProps();
                }
                else
                {
                    writer.ConfigVideo(videoStream.GetUncompressedVideoMediaType());
                }
                progressTracker.CurrentValue = 0;
				progressTracker.AVStatusMessage = "Transcoding AV";
			}

            //Now all the config and prep is done, so write the segment.
            writeSegment();

			//If there is a Presentation stream, process it here unless the slides were used for the video stream.
			if ((!useSlideStream) && (this.segment.PresentationDescriptor != null) &&
				(this.segment.PresentationDescriptor.PresentationCname != null))
			{
				progressTracker.CurrentValue = 0;
				progressTracker.ShowPercentComplete = false;
				progressTracker.AVStatusMessage = "Writing Presentation";
				/// The offset for PresentationMgr is a timespan in ticks to be subtracted from each absolute timestamp 
				/// to make a new absolute timestamp which has been adjusted for accumulted time skipped (or overlap) between 
				/// segments, or time during which there is missing AV data at the beginning of the first segment.
				///   It is calculated as: actualSegmentStart - jobStart - offset   
				///		where:
				///	actualSegmentStart: Time of the first AV write for the current segment
				///	JobStart: user requested start time for the first segment in this job (Presentation data reference time)
				///	offset: calculated actual duration of all segments previous to this one.
				///	
				///	Note that the presentation output will use the user-specified jobStart as the reference time.
				///	During playback, relative timings of presentation events will be calculated by subtracting the reference time.
				///	Also note that the reference time may not be the same as the actualSegmentStart for the first segment of the
				///	job in the case where there is missing data at the beginning.  This often (always) happens if we
				///	begin processing an archive from the beginning, since it takes several seconds to get the first AV bits
				///	into the database after the ArchiveService joins a venue.

				//long thisSegmentOffset = this.actualSegmentStart - startTime.Ticks;
				//long tmpoffset = this.actualSegmentStart-jobStart-offset;
				//Debug.WriteLine ("this segment offset. actualSegmentStart = " + this.actualSegmentStart.ToString() +
				//	" jobStart = " + jobStart.ToString() + " offset = " + offset.ToString() + 
				//	" offset to presenterMgr = " + tmpoffset.ToString());
                long previousSegmentEndTime = 0;
                if (m_PreviousSegment != null) {
                    previousSegmentEndTime = m_PreviousSegment.actualSegmentEnd;
                }
				presentationMgr = new PresentationMgr(this.segment.PresentationDescriptor, this.actualSegmentStart, this.actualSegmentEnd,
					this.actualSegmentStart-jobStart-offset,this.logMgr, previousSegmentEndTime, progressTracker);

				if (cancel)
					return null;
				String errmsg = presentationMgr.Process();
				if (errmsg !=null)
				{
					this.logMgr.WriteLine(errmsg);
					this.logMgr.ErrorLevel = 5;
				}
				progressTracker.ShowPercentComplete = true;
			}

            if ((useSlideStream) && (slideStream != null))
                slideStream.Dispose();

			return null;
		}

		/// <summary>
		/// If the segment contains presentation data, return the XML-formatted output.
		/// This is only valid after the Encode method has completed.
		/// </summary>
		/// <returns></returns>
		public String GetPresentationData()
		{
			if (presentationMgr != null)
				return presentationMgr.Print();
			else 
				return "";
		}

		/// <summary>
		/// Cancel an Encode method call in progress.
		/// </summary>
		public void Stop()
		{
			//Send stop signal to all potentially long-running processes.
			cancel = true;
			if (videoStream != null) videoStream.Stop();
            if (slideStream != null) slideStream.Stop();
			if (audioStreams != null)
			{
				for (int i=0;i<audioStreams.Length;i++)
				{
					if (audioStreams[i] != null)
						audioStreams[i].Stop();
				}
			}
			if (mixer != null) mixer.Stop();
			if (presentationMgr != null) 
				presentationMgr.Stop();
		}

		#endregion Public Methods

		#region Private Methods

        /// <summary>
        /// Read and write audio and video samples for the segment.
        /// </summary>
        private void writeSegment()
        {
            BufferChunk audioSample = null;
            BufferChunk videoSample = null;
            long videoTime = long.MaxValue;
            bool videoAvailable = true;
            long audioTime = long.MaxValue;
            bool audioAvailable = true;
            long refTime = 0;
            long lastWriteTime = 0;
            bool newstream = false;
            bool keyframe = false;
            bool discontinuity = true;
            bool junk1 = false;
            bool junk2 = false;
            progressTracker.EndValue = (int)((TimeSpan)(endTime - startTime)).TotalSeconds;


            while (!cancel)
            {
                if ((audioSample == null) && (audioAvailable))
                {
                    audioAvailable = getNextSample(PayloadType.dynamicAudio, out audioSample, out audioTime, out junk1, out junk2);
                }

                if ((videoSample == null) && (videoAvailable))
                {
                    videoAvailable = getNextSample(PayloadType.dynamicVideo, out videoSample, out videoTime, out newstream,out keyframe);
                }

                //If both audio and video are available, choose the one with the smaller timestamp
                if ((audioSample != null) && (videoSample != null))
                {
                    if (audioTime < videoTime)
                    {
                        //write audio
                        if (refTime == 0)
                            refTime = audioTime;
                        Debug.WriteLine("Write audio: " + (audioTime-refTime).ToString() + ";length=" + audioSample.Length.ToString());
                        lastWriteTime = audioTime - refTime;
                        writeSample(PayloadType.dynamicAudio, audioSample, (ulong)(audioTime - refTime + offset),newstream, keyframe, discontinuity);
                        audioSample = null;
                    }
                    else
                    {
                        //write video
                        if (refTime == 0)
                            refTime = videoTime;
                        Debug.WriteLine("Write video: " + (videoTime-refTime).ToString() + ";length=" + videoSample.Length.ToString());
                        lastWriteTime = videoTime - refTime;
                        writeSample(PayloadType.dynamicVideo, videoSample, (ulong)(videoTime - refTime + offset), newstream, keyframe, discontinuity);
                        videoSample = null;
                        discontinuity = false;
                    }
                }
                else if (audioSample != null) //write remaining audio
                {
                    if (refTime == 0)
                        refTime = audioTime;
                    Debug.WriteLine("Write audio: " + (audioTime-refTime).ToString() + ";length=" + audioSample.Length.ToString());
                    lastWriteTime = audioTime - refTime;
                    writeSample(PayloadType.dynamicAudio, audioSample, (ulong)(audioTime - refTime + offset), newstream, keyframe, discontinuity);
                    audioSample = null;
                }
                else if (videoSample != null) //write remaining video
                {
                    if (refTime == 0)
                        refTime = videoTime;
                    Debug.WriteLine("Write video: " + (videoTime-refTime).ToString() + ";length=" + videoSample.Length.ToString());
                    lastWriteTime = videoTime - refTime;
                    writeSample(PayloadType.dynamicVideo, videoSample, (ulong)(videoTime - refTime + offset), newstream, keyframe, discontinuity);
                    videoSample = null;
                    discontinuity = false;
                }
                else //done
                {
                    break;
                }
                progressTracker.CurrentValue = (int)(lastWriteTime / (Constants.TicksPerSec));
            }
            this.actualSegmentStart = refTime;
            this.actualSegmentEnd = lastWriteTime + refTime;        
        }

        /// <summary>
        /// Write the sample or stream sample in the correct way given the payload and the segment
        /// compression settings.  In the case of uncompressed video, also reconfigure the writer input properties
        /// if the newstream flag is set.  Keyframe and discontinuity flags are only relevant when writing
        /// compressed video.
        /// </summary>
        /// <param name="payload"></param>
        /// <param name="sample"></param>
        /// <param name="timestamp"></param>
        /// <param name="newstream"></param>
        /// <returns></returns>
        private void writeSample(PayloadType payload, BufferChunk sample, ulong timestamp, bool newstream, bool keyframe, bool discontinuity)
        {
            if (payload == PayloadType.dynamicAudio)
            {
                if (norecompression)
                {
                    writer.WriteCompressedAudio(timestamp, (uint)sample.Length, (byte[])sample);
                }
                else
                {
                    writer.WriteAudio((uint)sample.Length, sample, timestamp);
                }
            }
            else if (payload == PayloadType.dynamicVideo)
            { 
                if (norecompression)
                {
                    writer.WriteCompressedVideo(timestamp, (uint)sample.Length, (byte[])sample, keyframe, discontinuity);
                }
                else
                {
                    if (newstream)
                    {
                        //PRI3: don't need to reconfig unless MT changed.  Probably does no harm though
                        if (useSlideStream)
                        {
                            //writer.ConfigSlideStreamVideo(slideStream.UncompressedMediaType);
                            writer.ConfigVideo(slideStream.UncompressedMediaType);
                        }
                        else
                        {
                            writer.ConfigVideo(videoStream.GetUncompressedVideoMediaType());
                        }
                    }
                    writer.WriteVideo((uint)sample.Length, sample, timestamp);
                }
            }
        }


        /// <summary>
        /// Get the next sample from the appropriate source given the specified payload and the recompression and
        /// slide stream settings for this segment.
        /// </summary>
        private bool getNextSample(PayloadType payload, out BufferChunk sample, out long timestamp, out bool newstream, out bool keyframe)
        {
            if (payload == PayloadType.dynamicAudio)
            {
                if (norecompression)
                {
                    if (audiorecompression)
                    { 
                        newstream = false; keyframe= false;
                        return mixer.GetNextStreamSample(out sample, out timestamp);
                    }
                    else
                    {
                        keyframe = false;
                        return audioStreams[0].GetNextSample(out sample, out timestamp, out newstream);
                    }
                }
                else
                { 
                    newstream = false; keyframe = false;
                    return mixer.GetNextSample(out sample, out timestamp);
                }
            }
            else if (payload == PayloadType.dynamicVideo)
            {
                if (norecompression)
                { 
                    if (useSlideStream)
                    {
                        newstream = false;
                        //unsupported:
                        //return slideStream.GetNextStreamSample(out sample, out timestamp, out keyframe);
                    }
                    else
                    {
                        return videoStream.GetNextSample(out sample, out timestamp, out keyframe, out newstream);
                    }
                }
                else 
                { 
                    //Note: keyframe is not relevant for uncompressed video.
                    if (useSlideStream)
                    {
                        newstream = false; keyframe = false;
                        return slideStream.GetNextSample(out sample, out timestamp, out newstream);
                    }
                    else
                    {
                        keyframe = false;
                        return videoStream.GetNextSample(out sample, out timestamp, out newstream);
                    }
                }
            }
            sample = null;
            timestamp = 0;
            newstream = false;
            keyframe = false;
            return false;
        }

		#endregion Private Methods

        #region Test
        public void FrameGrabTest() {
            long start = DateTime.Parse(this.segment.StartTime).Ticks;
            long end = DateTime.Parse(this.segment.EndTime).Ticks;
            this.presentationMgr = new PresentationMgr(this.segment.PresentationDescriptor, start, end,
                    start - jobStart - offset, this.logMgr, 0, progressTracker);

            this.FrameGrabMgr.FrameGrabInterval = 1000;
            this.FrameGrabMgr.DifferenceThreshold = 250;

            this.presentationMgr.Process();
        }
        #endregion
    }
}
