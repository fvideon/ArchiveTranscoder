using System;
using System.Diagnostics;
using MSR.LST;
using MSR.LST.MDShow;
using MSR.LST.Net.Rtp;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Handle a collection of streams for one payload (audio or video only), one time range, one cname, 
	/// and an optional video camera name.  Determine Media types, build Windows Media profiles, and return 
	/// compressed or uncompressed stream data.
	/// </summary>
	/// 
	///	The streams may begin before, and end after the specified start/end times, and there may be gaps between them.  
	///	The StreamMgr hides these details from the caller and returns the data as one continuous set of audio or video samples.
	/// 
	/// The caller sets a flag to indicate whether samples returned should be compressed or uncompressed.
	/// 
	/// If compressed, GetNextSample returns samples and timestamp directly from the database.
	/// If uncompressed, the caller must use the 'ToRawWMFile' method before doing anything else. 
	/// The uncompressed samples and Mediatypes can then be returned to the caller.  
	public class StreamMgr
	{
		#region Members

		internal bool compressed;		//does the caller want compressed or uncompressed data?
		internal DBStreamPlayer[] streamPlayers;		//Database stream readers (compressed streams)
		internal FileStreamPlayer[] fileStreamPlayers;	//Read from files (uncompressed streams)
		internal int[] streams;			//The stream ID's at stake
		internal PayloadType payload;	//Audio or video payload
		internal bool cancel;			//Cancel long-running method call
		private string cname;			//Participant cname
		private int lastFSP;			//Index of Most recently read FileStreamPlayer
		private Guid currentFSPGuid;	//Guid of most recently read FileStreamPlayer
		private string name;			//Optional camera name for video sources
        private ProfileData streamProfileData;  //Compressed mediaType and codec private data for this stream.
        private long lookBehindStart;       //When reading compressed data we want to look behind to be sure to get all the data at the start.
        private long startTime;             //The requested start time.
        private int currentChannels;
		#endregion Members

		#region Properties

		/// <summary>
		/// Identifier for the most recently read stream (uncompressed)
		/// </summary>
		/// This is used to enable the Audio mixer to exclude data from incompatible streams.
		public Guid CurrentFSPGuid
		{
			get { return currentFSPGuid; }
		}


        public int CurrentChannels {
            get { return currentChannels; }
        }

        /// <summary>
        /// ProfileData containing compressed MediaType and codec private data for this stream.
        /// This is set only after a call to ToRawWMFile();
        /// </summary>
        public ProfileData StreamProfileData
        {
            get { return streamProfileData; }
        }

		#endregion Properties

		#region Ctor

		public StreamMgr(String cname, String name, DateTime start, DateTime end, bool compressed, PayloadType payload)
		{
            streamProfileData = null;
			cancel = false;
			this.compressed = compressed;
			lastFSP = -1;
			currentFSPGuid = Guid.Empty;
		
			this.cname = cname;
			this.name = name;
			this.payload = payload;

            //Look behind to make sure we get a key frame before the requested start time.
            //Otherwise we may have several seconds of missing video frames at the beginning of the segment.
            //PRI2: make the lookbehind customizable.
            //PRI2: audio surely doesn't need as much lookbehind?  If any?
            startTime = start.Ticks;
            lookBehindStart = startTime;
            if (!compressed)
                lookBehindStart -= Constants.TicksPerSec * 10;

			//Get the relevant stream_id's and create DBStreamPlayers for each.
			streams = DatabaseUtility.GetStreams(payload,cname,name,startTime,end.Ticks);
			streamPlayers = new DBStreamPlayer[streams.Length];
			for(int i=0; i<streams.Length; i++)
			{
                streamPlayers[i] = new DBStreamPlayer(streams[i],lookBehindStart,end.Ticks,payload);
            }
		}

		#endregion Ctor

		#region Public Methods

		/// <summary>
		/// Stop a currently running invocation of ToRawWMFile.
		/// </summary>
		public void Stop()
		{
			cancel = true;
		}

		/// <summary>
		/// Return the next sample, compressed or uncompressed as appropriate.
		/// </summary>
		/// <param name="sample">The bits</param>
		/// <param name="timestamp">Absolute timestamp in ticks</param>
		/// <param name="newstream">True if this sample is uncompressed and
		/// is from a different stream than the previous sample (thus may have a different media type).</param>
		/// <returns>True if a valid sample is returned</returns>
		public bool GetNextSample(out BufferChunk sample, out long timestamp, out bool newstream)
		{
			bool keyframe = false;
			return GetNextSample(out sample, out timestamp, out keyframe, out newstream);
		}
		
		/// <summary>
		/// Return the next sample, compressed or uncompressed as appropriate.
		/// </summary>
		/// <param name="sample">The bits</param>
		/// <param name="timestamp">Absolute timestamp in ticks</param>
		/// <param name="keyframe">True if this sample is compressed and is a video keyframe</param>
		/// <param name="newstream">True if this sample is uncompressed and
		/// is from a different stream than the previous sample (thus may have a different media type).</param>
		/// <returns>True if a valid sample is returned</returns>
		public bool GetNextSample(out BufferChunk sample, out long timestamp, out bool keyframe, out bool newstream)
		{
			BufferChunk frame;
			keyframe = false;
			newstream = false;
			if (compressed)
			{
				for (int i=0;i<streamPlayers.Length;i++)
				{
					if (streamPlayers[i].GetNextFrame(out frame,out timestamp))
					{
						sample = ProfileUtility.FrameToSample(frame, out keyframe);
						return true;
					}
				}

			}
			else
			{
				for (int i=0;i<fileStreamPlayers.Length;i++)
				{
					while (fileStreamPlayers[i].GetNextSample(out sample, out timestamp))
					{
                        if (timestamp >= startTime) //skip past frames that may be returned due to the look behind.
                        {
                            currentFSPGuid = fileStreamPlayers[i].xGuid;
                            if (fileStreamPlayers[i].AudioMediaType != null) {
                                currentChannels = fileStreamPlayers[i].AudioMediaType.WaveFormatEx.Channels;
                            }
                            //PRI3: instead of newstream, we could track the MT, and raise a flag only when it changes.
                            if (i != lastFSP)
                            {
                                newstream = true;
                            }
                            lastFSP = i;
                            return true;
                        }
					}
				}
			}

			sample = null;
			timestamp = 0;
			return false;
		}


		/// <summary>
		/// Make sure all the streams for this StreamMgr have compatible compressed media types.
		/// Also check against the media type, optionally supplied as an argument.
		/// Log details about any problems we find.
		/// </summary>
		/// In no-recompression scenarios including the preview we check this and warn the user, 
		/// though we can't do anything about it.
		/// <param name="prevMT">An additional MediaType to check against, or null if none</param>
		/// <param name="log">Where to write details about any problems found</param>
		/// <returns>always true</returns>
		public bool ValidateCompressedMT(MediaType prevMT, LogMgr log)
		{

			MediaType lastMT = prevMT;
			for(int i=0; i<streamPlayers.Length; i++)
			{
				if (payload == PayloadType.dynamicVideo)
				{
					if (!ProfileUtility.CompareVideoMediaTypes((MediaTypeVideoInfo)streamPlayers[i].StreamMediaType,(MediaTypeVideoInfo)lastMT))
					{
						Debug.WriteLine("incompatible video mediatype found.");
						log.WriteLine("Warning: A change in the media type was found in the video stream from " + this.cname +
							" beginning at " + streamPlayers[i].Start.ToString() + ".  Without recompression, this may cause " +
							" a problem with the output.  Using recompression should resolve the issue. ");
						log.ErrorLevel = 5;
					}
				}
				else if (payload == PayloadType.dynamicAudio)
				{
					if (!ProfileUtility.CompareAudioMediaTypes((MediaTypeWaveFormatEx)streamPlayers[i].StreamMediaType,(MediaTypeWaveFormatEx)lastMT))
					{
						Debug.WriteLine("incompatible mediatype found.");
						log.WriteLine("Warning: A change in the media type was found in the audio stream from " + this.cname +
							" beginning at " + streamPlayers[i].Start.ToString() + ".  Without recompression, this may cause " +
							" a problem with the output.  Using recompression should resolve the issue. ");
						log.ErrorLevel = 5;
					}
				}
				lastMT = streamPlayers[i].StreamMediaType;
			}
			return true;
		}


		/// <summary>
		/// Get the timestamp of the compressed sample to be return by the next call to GetNextSample.
		/// Return false if there are no more, or if we are configured to return uncompressed samples.
		/// </summary>
		/// <param name="timestamp">Next timestamp (absolute time in ticks)</param>
		/// <returns>true there is another compressed sample to read</returns>
		public bool GetNextSampleTime(out long timestamp)
		{
			if (compressed)
			{
				for (int i=0;i<streamPlayers.Length;i++)
				{
					if (streamPlayers[i].GetNextFrameTime(out timestamp))
					{
						return true;
					}
				}
			}

			timestamp = 0;
			return false;
		}

		/// <summary>
		/// Write each stream from DBStreamPlayer to a WM file, then create FileStreamPlayers for each.
		/// It is necessary to do this before reading uncompressed samples, or using any of the 
		/// methods that return uncompressed MediaTypes.
		/// This can be a long-running process.  It can be cancelled with the Stop method.
		/// </summary>
		/// <returns>False if we failed to configure the native profile</returns>
		public bool ToRawWMFile(ProgressTracker progressTracker)
		{
			if (cancel)
				return true;

			String tmpfile = "";
			fileStreamPlayers = new FileStreamPlayer[streamPlayers.Length];
			for (int i=0;i<streams.Length;i++)
			{
                streamProfileData = ProfileUtility.StreamIdToProfileData(streams[i],payload);
				if (payload==PayloadType.dynamicVideo)
				{
					tmpfile = Utility.GetTempFilePath("wmv");
					//nativeProfile = ProfileUtility.MakeNativeVideoProfile(streams[i]);
				}
				else
				{
					tmpfile = Utility.GetTempFilePath("wma");
					//nativeProfile = ProfileUtility.MakeNativeAudioProfile(streams[i]);
				}
				WMWriter wmWriter = new WMWriter();
				wmWriter.Init();
				//if (!wmWriter.ConfigProfile(nativeProfile,"",0))
                if (!wmWriter.ConfigProfile(StreamProfileData))
				{
					return false;
				}
				wmWriter.ConfigFile(tmpfile);
				wmWriter.ConfigNullProps();
				//wmWriter.SetCodecInfo(payload);
				wmWriter.Start();

				long streamTime=long.MaxValue;
				long refTime=0;
				long endTime=0;
				long lastWriteTime=0;
				BufferChunk frame;
				BufferChunk sample;
				bool keyframe;
				bool discontinuity;
				discontinuity = true;
                //Catch exceptions to work around the rare case of data corruption.
                //Oddly in one case where this occurred it did not occur if the segments were short enough
                while (streamPlayers[i].GetNextFrame(out frame, out streamTime)) {
                    try {
                        sample = ProfileUtility.FrameToSample(frame, out keyframe);
                    }
                    catch {
                        DateTime dt = new DateTime(streamTime);
                        Console.WriteLine("Possible data corruption in stream: " + this.payload + ";" + this.cname +
                            " at " + dt.ToString() + " (" + streamTime.ToString() + ")");
                        continue;
                    }
                    if (refTime == 0)
                        refTime = streamTime;
                    lastWriteTime = streamTime - refTime;
                    try {
                        if (payload == PayloadType.dynamicVideo) {
                            //Debug.WriteLine("Write video: " + (streamTime-refTime).ToString() + ";length=" + sample.Length.ToString());
                            wmWriter.WriteCompressedVideo((ulong)(streamTime - refTime), (uint)sample.Length, (byte[])sample, keyframe, discontinuity);
                        }
                        else {
                            //Debug.WriteLine("Write audio: " + (streamTime-refTime).ToString() + ";length=" + sample.Length.ToString());
                            wmWriter.WriteCompressedAudio((ulong)(streamTime - refTime), (uint)sample.Length, (byte[])sample);
                        }
                    }
                    catch {
                        DateTime dt = new DateTime(streamTime);
                        Console.WriteLine("Failed to write.  Possible data corruption in stream: " + this.payload + ";" + this.cname +
                            " at " + dt.ToString() + " (" + streamTime.ToString() + ")");
                    } 
                    
                    if (discontinuity)
                        discontinuity = false;
                    endTime = streamTime;
                    if (cancel)
                        break;

                    progressTracker.CurrentValue = (int)(lastWriteTime / Constants.TicksPerSec);
                    //Debug.WriteLine("StreamMgr.ToRawWMFile: ProgressTracker currentValue=" + progressTracker.CurrentValue.ToString() +
                    //    ";streamTime=" + streamTime.ToString());

                }

				wmWriter.Stop();
				wmWriter.Cleanup();
				wmWriter=null;

				fileStreamPlayers[i] = new FileStreamPlayer(tmpfile,refTime,endTime,false,streams[i]);
				if (cancel)
					break;
			}
			return true;
		}

		/// <summary>
		/// Return the most recently read uncompressed video media type.  If no uncompressed video samples 
		/// have been returned yet, return the type for the first sample.
		/// This is only valid after the method 'ToRawWMFile' has completed.
		/// </summary>
		/// <returns></returns>
		public MediaTypeVideoInfo GetUncompressedVideoMediaType()
		{
			if (payload==PayloadType.dynamicVideo)
			{
				if (fileStreamPlayers != null)
				{
					if (fileStreamPlayers.Length>0)
					{
						if (lastFSP >= 0)
							return fileStreamPlayers[lastFSP].VideoMediaType;
						else
							return fileStreamPlayers[0].VideoMediaType;
					}
				}
			}
			return null;
		}

		/// <summary>
		/// Return the MediaType of the first uncompressed audio sample.
		/// This is only valid after the method 'ToRawWMFile' has completed.
		/// </summary>
		/// <returns></returns>
		public MediaTypeWaveFormatEx GetUncompressedAudioMediaType()
		{
			if (payload==PayloadType.dynamicAudio)
			{
				if (fileStreamPlayers != null)
				{
					if (fileStreamPlayers.Length>0)
					{
						return fileStreamPlayers[0].AudioMediaType;
					}
				}
			}
			return null;
		}


		/// <summary>
		/// Given a AudioCompatibilityMgr, feed it all the uncompressed audio streams known to this StreamMgr.
		/// This is only valid after the method 'ToRawWMFile' has completed.
		/// </summary>
		/// <param name="audioCompatibilityMgr"></param>
		public void CheckUncompressedAudioTypes(AudioCompatibilityMgr audioCompatibilityMgr)
		{
			if (payload != PayloadType.dynamicAudio)
				return;

			if (fileStreamPlayers != null)
			{
				foreach (FileStreamPlayer fsp in fileStreamPlayers)
				{
					audioCompatibilityMgr.Check(fsp.AudioMediaType,fsp.Duration,fsp.xGuid,this.cname,this.name,fsp.StartTime,fsp.StreamID);
				}
			}
		}


		/// <summary>
		/// Retrieve the native compressed video media type for the last video stream known to this StreamMgr. 
		/// </summary>
		/// <returns></returns>
		public MediaTypeVideoInfo GetFinalCompressedVideoMediaType()
		{
			if (payload==PayloadType.dynamicVideo)
			{
				if (streamPlayers != null)
				{
					if (streamPlayers.Length>0)
					{	
						return (MediaTypeVideoInfo)streamPlayers[streamPlayers.Length-1].StreamMediaType;
					}
				}
			}
			return null;
		}

		/// <summary>
		/// Retrieve the native compressed audio media type for the last audio stream known to this StreamMgr. 
		/// </summary>
		/// <returns></returns>
		public MediaTypeWaveFormatEx GetFinalCompressedAudioMediaType()
		{
			if (payload==PayloadType.dynamicAudio)
			{
				if (streamPlayers != null)
				{
					if (streamPlayers.Length>0)
					{
						
						return (MediaTypeWaveFormatEx)streamPlayers[streamPlayers.Length-1].StreamMediaType;
					}
				}
			}
			return null;
		}

		#endregion Public Methods
	}
}
