using System;
using System.Collections;
using System.Diagnostics;
using MSR.LST;
using MSR.LST.MDShow;
using System.Text;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Basic PCM Audio Mixer.
	/// </summary>
	public class AudioMixer
	{
		#region Members

		private uint bitsPerSample;
		private uint bytesPerSample;
        private uint ticksPerSample;
		private long limit;
		private StreamMgr[] audioMgr;
		private SampleBuffer[] buffers;
		private long refTime;
		private ArrayList inbufs;
		private bool cancel;
		private FileStreamPlayer fileStreamPlayer;
		private MediaTypeWaveFormatEx uncompressedMediaType;
		private LogMgr log;
		private ArrayList incompatibleGuids;
		private int compatibleStreamID;

		#endregion Members

		#region Properties

		/// <summary>
		/// The uncompressed Media Type for use when configuring the WM Writer.
		/// </summary>
		public MediaTypeWaveFormatEx UncompressedAudioMediaType
		{
			get { return uncompressedMediaType; }
		}

		#endregion Properties

		#region Ctor

        /// <summary>
        /// Assume that all the StreamMgrs provided are configured with the correct time range. Do not assume 
        /// contiguous data, or compatible media types.
        /// </summary>
        /// <param name="audioMgr"></param>
        /// <param name="log"></param>
		public AudioMixer(StreamMgr[] audioMgr, LogMgr log)
		{
			fileStreamPlayer = null;
			this.audioMgr = audioMgr;
			this.log = log;
			refTime = long.MinValue;
			inbufs = new ArrayList();

			if (audioMgr.Length==0)
			{
				return;
			}
			

			/// Examine the uncompressed MT's for each audioMgr, and implement a voting system so that the 
			/// media type that is dominant for this mixer is the one we use, and other incompatible MT's 
			/// are ignored in the mix.  Log a warning at places where the MT changes.
			/// Remember that each audioMgr may itself have multiple FileStreamPlayers which have different uncompressed
			/// media types.  
			/// Finally we need to make our uncompressed MT available to the caller for use in configuring the writer.
			/// Later on let's look at ways to convert any common uncompressed types so that they are compatible.
			AudioCompatibilityMgr audioCompatibilityMgr = new AudioCompatibilityMgr();

			foreach(StreamMgr astream in audioMgr)
			{
				astream.CheckUncompressedAudioTypes(audioCompatibilityMgr);
			}

			this.uncompressedMediaType = audioCompatibilityMgr.GetDominantType();
			String warning = audioCompatibilityMgr.GetWarningString();
			if (warning != "")
			{
				log.WriteLine(warning);
				log.ErrorLevel = 5;
			}
			incompatibleGuids = audioCompatibilityMgr.GetIncompatibleGuids();

			// Here we also want to collect a "native" (compressed) profile corresponding to one of the "compatible" 
			// streams.  This is useful in case we need to recompress.  Note this profile can be created if we have 
			// a stream ID.
			compatibleStreamID = audioCompatibilityMgr.GetCompatibleStreamID();

			this.bitsPerSample = this.uncompressedMediaType.WaveFormatEx.BitsPerSample;
			this.bytesPerSample = bitsPerSample/8;          
            this.ticksPerSample = ((uint)Constants.TicksPerSec)/(this.uncompressedMediaType.WaveFormatEx.SamplesPerSec * this.uncompressedMediaType.WaveFormatEx.Channels);
			limit = (long)((ulong)1 << (int)bitsPerSample) / 2 - 1; //clip level

            buffers = new SampleBuffer[audioMgr.Length];
            for (int i = 0; i < buffers.Length; i++) {
                buffers[i] = new SampleBuffer(audioMgr[i],ticksPerSample,incompatibleGuids,this.uncompressedMediaType.WaveFormatEx.Channels);
            }

		}

		#endregion Ctor

		#region Public Methods

        public bool GetNextSample(out BufferChunk sample, out long time) {
            sample = null;
            const int SAMPLE_SIZE = 4096;

            sample = Mix(SAMPLE_SIZE, out time);
            
            if (refTime == long.MinValue) {
                refTime = time;
            }
            
            if (sample == null)
                return false;

            return true;
        }

        /// <summary>
		/// Get the next compressed sample. Note that calls to this method are only valid after a 
		/// successful call of the Recompress method.
		/// </summary>
		/// <param name="sample"></param>
		/// <param name="time"></param>
		/// <returns></returns>
		public bool GetNextStreamSample(out BufferChunk sample, out long time)
		{
			sample = null;
			time = 0;

			if (fileStreamPlayer == null)
			{
				return false;
			}
			bool ret = fileStreamPlayer.GetNextSample(out sample, out time);
			return ret;
		}

		/// <summary>
		/// Recompress audio from mixer into a temp file using the native profile.  This is used to implement mixing
		/// in the 'norecompression' scenario.  
		/// </summary>
		/// <param name="progressTracker"></param>
		/// <returns></returns>
		public bool Recompress(ProgressTracker progressTracker)
		{
			cancel = false;

			if (audioMgr.Length == 0)
				return false;

			//String useProfile;
            ProfileData profileData = null;

			if (this.compatibleStreamID>=0)
			{
                profileData = ProfileUtility.StreamIdToProfileData(compatibleStreamID, MSR.LST.Net.Rtp.PayloadType.dynamicAudio);
				//Debug.WriteLine("Mixer.Recompress: using audio profile from streamID: " + compatibleStreamID.ToString());
			}
			else
			{
				//Under what circumstances could we get here??
                profileData = audioMgr[0].StreamProfileData;
			}

			WMWriter wmWriter = new WMWriter();
			wmWriter.Init();

            if (!wmWriter.ConfigProfile(profileData))
			{
				return false;
			}

			String tempFileName = Utility.GetTempFilePath("wma");
			wmWriter.ConfigFile(tempFileName);
			wmWriter.GetInputProps();
			wmWriter.ConfigAudio(audioMgr[0].GetUncompressedAudioMediaType());

			wmWriter.Start();

			//Write samples
			progressTracker.CurrentValue = 0;
			BufferChunk audioSample = null;
			long audioTime=long.MaxValue;
			long refTime=0, endTime=0;
			long lastWriteTime=0;

			while(!cancel)
			{
				if (audioSample==null)
				{
					endTime = audioTime;
					if (!GetNextSample(out audioSample, out audioTime))
					{
						break;
					}
				}

				if (audioSample!=null)
				{
					//write audio
					if (refTime==0)
						refTime = audioTime;
					//Debug.WriteLine("mixer.Recompress write audio: " + (audioTime-refTime).ToString() + ";length=" + audioSample.Length.ToString());
					lastWriteTime=audioTime-refTime;
					wmWriter.WriteAudio((uint)audioSample.Length,audioSample,(ulong)(audioTime-refTime));
					audioSample=null;
				}
				else
				{
					break;
				}	
				progressTracker.CurrentValue = (int)(lastWriteTime/(Constants.TicksPerSec));
			}

			wmWriter.Stop();
			wmWriter.Cleanup();
			wmWriter=null;
			
			//Prepare a filestreamPlayer to read back compressed samples.
			fileStreamPlayer = new FileStreamPlayer(tempFileName,refTime,endTime,true, -1);

			return true;
		}

		/// <summary>
		/// Stop a currently running invocation of Recompress
		/// </summary>
		public void Stop()
		{
			cancel = true;
		}

		#endregion Public Methods

		#region Private Methods

        /// <summary>
        /// Mix the bits
        /// </summary>
        /// <param name="inbufs"></param>
        /// <returns></returns>
        private BufferChunk Mix(int sampleSize, out long timestamp) {

            int sampleCount = (int)(sampleSize / bytesPerSample);
            byte[] outBuf = new byte[sampleSize];
            long MixedSample;

            //find minimum timestamp across all buffers
            long minTimestamp = long.MaxValue;
            for (int i = 0; i < buffers.Length; i++) {
                if (!buffers[i].Exhausted) {
                    long nextTs = buffers[i].GetNextTimestamp();
                    if (nextTs < minTimestamp)
                        minTimestamp = nextTs;
                }
            }

            timestamp = minTimestamp;

            //return null if all the streams are exhausted.
            if (minTimestamp == long.MaxValue)
                return null;

            for (int i = 0; i < sampleCount; i++) {
                int offset = (int)(i * bytesPerSample);

                MixedSample = 0;
                foreach (SampleBuffer sb in buffers) {
                    MixedSample+=sb.ReadSample((int)this.bytesPerSample,minTimestamp);
                }
                minTimestamp += ticksPerSample;

                //Do limiting
                if (MixedSample > limit) MixedSample = limit;
                if (MixedSample < -limit) MixedSample = -limit;

                //Convert the sample back to a byte[]and put it into the buffer
                //Note: as of CXP4 we should only need to worry about 8 and 16 bits per sample.
                switch (bitsPerSample) {
                    case 8:
                        outBuf[offset] = (byte)MixedSample;
                        break;
                    case 16:
                        outBuf[offset + 1] = (byte)(MixedSample >> 8);
                        outBuf[offset] = (byte)(MixedSample);
                        break;
                    case 32:
                        outBuf[offset + 3] = (byte)(MixedSample >> 24);
                        outBuf[offset + 2] = (byte)(MixedSample >> 16);
                        outBuf[offset + 1] = (byte)(MixedSample >> 8);
                        outBuf[offset] = (byte)(MixedSample);
                        break;
                    default:
                        break;
                }
            }

            return new BufferChunk(outBuf);
        }

        #endregion Private Methods

        #region SampleBuffer Class

        /// <summary>
		/// Buffer management to aid in audio mixing.  Feed samples and timestamps to the mixer.  
        /// Set a flag when there are no more samples available.
		/// </summary>
		private class SampleBuffer
		{
			private BufferChunk buffer;
            private StreamMgr streamMgr;
            private long nextTimestamp;
            private uint ticksPerSample;
            private ArrayList incompatibleGuids;
            private bool exhausted;
            private int targetChannels;
            private int currentChannels;
            private long previousSample;
            private int sampleCounter;

            //Add our sample to the mix if its timestamp is within this range:
            private int FUDGE_FACTOR = Constants.TicksPerMs * 100;

            public bool Exhausted {
                get { return exhausted;  }
            }

            public SampleBuffer(StreamMgr streamMgr,uint ticksPerSample, ArrayList incompatibleGuids, int targetChannels) {
                this.streamMgr = streamMgr;
                this.ticksPerSample = ticksPerSample;
                this.incompatibleGuids = incompatibleGuids;
                this.targetChannels = targetChannels;
                sampleCounter = 0;
                exhausted = false;
            }

            /// <summary>
            /// Read the next sample at the given timestamp (or within the FUDGE_FACTOR range.)
            /// Return zero if there is no sample available for this timestamp.  
            /// </summary>
            /// <param name="bitsPerSample"></param>
            /// <returns></returns>
            public long ReadSample(int bytesPerSample, long timestamp) {
                if ((buffer == null) || (buffer.Length < bytesPerSample)) {
                    if ((buffer != null) && (buffer.Length > 0)) {
                        //Debug.WriteLine(".");
                        Debug.WriteLine("Warning: An audio chunk appeared to contain a fractional part of a sample which will be discarded.");
                    }
                    FillBuffer();
                }

                if (exhausted)
                    return 0;

                bool convertToStereo = false;
                bool convertToMono = false;

                if (currentChannels != targetChannels) {
                    if (targetChannels == 2) {
                        convertToStereo = true;
                    }
                    else {
                        convertToMono = true;
                    }
                }

                long sampleAsLong = 0;
                if (Math.Abs(nextTimestamp - timestamp) < FUDGE_FACTOR) {
                    nextTimestamp += this.ticksPerSample;
                    if ((convertToStereo) && (sampleCounter % 2 == 1)) {
                        sampleAsLong = previousSample;
                    }
                    else {
                        switch (bytesPerSample) {
                            case 1:
                                sampleAsLong = buffer.NextByte();
                                if (convertToMono) { 
                                    sampleAsLong += buffer.NextByte();
                                    sampleAsLong = sampleAsLong / 2;                                
                                }
                                break;
                            case 2:
                                sampleAsLong = BitConverter.ToInt16(buffer.Buffer, buffer.Index);
                                //Note that NextInt16, NextInt32 assume the wrong byte order.  We just use them to 
                                //update the BufferChunk.Index and Length.
                                buffer.NextInt16();
                                if (convertToMono) { 
                                    sampleAsLong += BitConverter.ToInt16(buffer.Buffer, buffer.Index);
                                    buffer.NextInt16();
                                    sampleAsLong = sampleAsLong / 2;                                
                                }
                                break;
                            case 4:
                                sampleAsLong = BitConverter.ToInt32(buffer.Buffer, buffer.Index);
                                buffer.NextInt32();
                                if (convertToMono) { 
                                    sampleAsLong += BitConverter.ToInt32(buffer.Buffer, buffer.Index);
                                    buffer.NextInt32();
                                    sampleAsLong = sampleAsLong / 2;            
                                }
                                break;
                            default:
                                break;
                        }
                    }
                }

                this.previousSample = sampleAsLong;
                this.sampleCounter++;
                return sampleAsLong;
            }

            /// <summary>
            /// Return the timestamp of the next available sample.  If the stream is exhausted, return
            /// long.MaxValue.
            /// </summary>
            /// <returns></returns>
            public long GetNextTimestamp() {
                if ((buffer == null) || (buffer.Length == 0)) {
                    FillBuffer();
                }

                if (exhausted)
                    return long.MaxValue;

                return nextTimestamp;
            }

            /// <summary>
            /// Get the next sample from the stream manager.  Anything remaining in the current buffer is discarded.
            /// Update the timestamp.  Skip over samples from streams that are known to be of incompatible media types.
            /// </summary>
            private void FillBuffer() {
                if (exhausted)
                    return;

                bool newstream;
                if (this.streamMgr.GetNextSample(out this.buffer, out this.nextTimestamp, out newstream)) {
                    this.currentChannels = streamMgr.CurrentChannels;
                    if (this.incompatibleGuids.Contains(streamMgr.CurrentFSPGuid)) {
                        while (true) {
                            if (this.streamMgr.GetNextSample(out this.buffer, out this.nextTimestamp, out newstream)) {
                                if (!this.incompatibleGuids.Contains(streamMgr.CurrentFSPGuid)) {
                                    break;
                                }
                            }
                            else {
                                exhausted = true;
                                break;
                            }
                        }
                    }
                }
                else {
                    exhausted = true;
                }
            }


		}


		#endregion SampleBuffer Class
	}

}


