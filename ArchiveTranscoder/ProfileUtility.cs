using System;
using System.Diagnostics;
using System.Text;
using MSR.LST.Net.Rtp;
using MSR.LST;
using MSR.LST.MDShow;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Static methods to help building Windows Media profiles and MediaTypes from raw frames, and 
	/// from byte arrays copied from WM Interop.
	/// </summary>
	public class ProfileUtility
	{
		public ProfileUtility(){}

		#region Public Static

        /// <summary>
		/// Construct and return a ProfileData instance
        /// containing media type and codec private data for audio and video as
        /// determined using the first audio and video frames referenced by the segment.
        /// Audio sources in the segment other than the first will be ignored.
        /// </summary>
        /// <param name="segment"></param>
        /// <returns></returns>
        public static ProfileData SegmentToProfileData(ArchiveTranscoderJobSegment segment)
        {
            DateTime startDt = DateTime.Parse(segment.StartTime);
            DateTime endDt = DateTime.Parse(segment.EndTime);

            byte[] aframe = DatabaseUtility.GetFirstFrame(PayloadType.dynamicAudio, segment.AudioDescriptor[0].AudioCname,
                segment.AudioDescriptor[0].AudioName, startDt.Ticks, endDt.Ticks);

            if (Utility.SegmentFlagIsSet(segment, SegmentFlags.SlidesReplaceVideo))
            {
                return AudioFrameToProfileData(aframe);
            }
            else
            {
                byte[] vframe = DatabaseUtility.GetFirstFrame(PayloadType.dynamicVideo, segment.VideoDescriptor.VideoCname,
                    segment.VideoDescriptor.VideoName, startDt.Ticks, endDt.Ticks);
                return FramesToProfileData(aframe, vframe);
            }
        }

        /// <summary>
        /// Return a ProfileData instance containing the MediaType and codec private data as
        /// determined by the first frame of the stream referenced by the streamID.
        /// Return null if there are no frames, or if stream payload is not audio or video.
        /// </summary>
        /// <param name="streamID"></param>
        /// <returns></returns>
        public static ProfileData StreamIdToProfileData(int streamID, PayloadType payload)
        { 			
            byte[] frame = DatabaseUtility.GetFirstFrame(streamID);
			if (frame==null)
				return null;

            if (payload == PayloadType.dynamicAudio)
            {
                return FramesToProfileData(frame, null);
            }
            else if (payload == PayloadType.dynamicVideo)
            {
                return FramesToProfileData(null, frame);
            }
            return null;
        }


		/// <summary>
		/// Given a Stream ID for an audio stream, return the Audio MediaType.
		/// </summary>
		/// <param name="id"></param>
		/// <returns></returns>
		public static MediaTypeWaveFormatEx StreamIdToAudioMediaType(int id)
		{
			byte[] frame = DatabaseUtility.GetFirstFrame(id);
			if (frame==null)
				return null;
			byte[] compressionData;
			return AudioMediaTypeFromFrame(frame, out compressionData);
		}

		/// <summary>
		/// Given a Stream ID for a video stream, return the video MediaType.
		/// </summary>
		/// <param name="id"></param>
		/// <returns></returns>
		public static MediaTypeVideoInfo StreamIdToVideoMediaType(int id)
		{
			byte[] frame = DatabaseUtility.GetFirstFrame(id);
			if (frame==null)
				return null;
			byte[] compressionData;
			return VideoMediaTypeFromFrame(frame, out compressionData);
		}

		/// <summary>
		/// Compare important fields to make sure the Audio MediaTypes are "compatible".  This is used
		/// with compressed audio to make sure we won't cause a Windows Media Writer object to 
		/// except when we feed it stream samples from multiple RTP audio streams.  
		/// By definition a null is compatible with any media type.
		/// Note: We do something similar for uncompressed samples in AudioTypeMonitor to validate MediaTypes 
		/// prior to audio mixing.
		/// </summary>
		/// PRI2: It is not totally clear that we check the correct set of parameters to cover all cases.
		/// <param name="mt1"></param>
		/// <param name="mt2"></param>
		/// <returns></returns>
		public static bool CompareAudioMediaTypes(MediaTypeWaveFormatEx mt1, MediaTypeWaveFormatEx mt2)
		{
			if ((mt1==null) || (mt2==null))
				return true;

			if ((mt1.MajorType != mt2.MajorType) ||
				(mt1.SubType != mt2.SubType) ||
				(mt1.FormatType != mt2.FormatType))
			{
				return false;
			}

			if ((mt1.WaveFormatEx.SamplesPerSec != mt2.WaveFormatEx.SamplesPerSec))
			{
				return false;
			}

			return true;
		}

		/// <summary>
		/// Compare important fields to make sure the MediaTypes are "compatible".  
		/// This is used with compressed samples to make sure we won't cause the Windows Media Writer
		/// object to except by feeding it stream samples from multiple RTP streams.
		/// By definition a null is compatible with any media type.
		/// Note: in the case of uncompressed samples, we can reconfigure the WM Writer to accept the 
		/// new media type, so there is no need to do this checking.
		/// </summary>
		/// <param name="mt1"></param>
		/// <param name="mt2"></param>
		/// <returns></returns>
		public static bool CompareVideoMediaTypes(MediaTypeVideoInfo mt1, MediaTypeVideoInfo mt2)
		{
			if ((mt1==null) || (mt2==null))
				return true;

			if ((mt1.MajorType != mt2.MajorType) ||
				(mt1.SubType != mt2.SubType) ||
				(mt1.FormatType != mt2.FormatType))
			{
				return false;
			}

			if ((mt1.VideoInfo.BitRate != mt2.VideoInfo.BitRate))
			{
				return false;
			}

			if ((mt1.VideoInfo.BitmapInfo.Height != mt2.VideoInfo.BitmapInfo.Height) ||
				(mt1.VideoInfo.BitmapInfo.Width != mt2.VideoInfo.BitmapInfo.Width) ||
				(mt1.VideoInfo.BitmapInfo.SizeImage != mt2.VideoInfo.BitmapInfo.SizeImage))
			{
				return false;
			}

            //special case for the Screen Streaming codec we had to repair these two members, so
            // only test them if not MSS2.
            if ((mt1.SubType != SubType.MSS2) &&
                ((mt1.VideoInfo.BitmapInfo.Compression != mt2.VideoInfo.BitmapInfo.Compression) ||
				(mt1.VideoInfo.BitmapInfo.BitCount != mt2.VideoInfo.BitmapInfo.BitCount)))
            {
                return false;
            }

			return true;
		}

		/// <summary>
		/// Given a byte[] (in the form of a BufferChunk) containing a MediaType received from WM Interop, and an
		/// allocated instance of MediaType, fill in the MediaType from the byte[].  This is just the 
		/// payload-independant part.
		/// </summary>
		/// <param name="mt"></param>
		/// <param name="bc"></param>
		public static void ReconstituteBaseMediaType(MediaType mt, BufferChunk bc)
		{
			Guid g = NextGuid(bc);
			mt.MajorType = MediaType.MajorTypeGuidToEnum(g);
			mt.mt.majortype = g;
			g = NextGuid(bc);
			mt.SubType = MediaType.SubTypeGuidToEnum(g);
			mt.mt.subtype = g;
			mt.FixedSizeSamples = (NextInt32(bc) == 0)?false:true;
			mt.TemporalCompression = (NextInt32(bc) == 0)?false:true;
			mt.SampleSize = NextInt32(bc);
			g = NextGuid(bc);
			mt.FormatType = MediaType.FormatTypeGuidToEnum(g);
			mt.mt.formattype = g;
			IntPtr punk = (IntPtr)NextInt32(bc);
			int cbformat = NextInt32(bc);
			IntPtr pbformat = (IntPtr)NextInt32(bc); //This not a valid pointer in this context.
		}

		/// <summary>
		/// Fill in the audio-specific parts of the MediaType from the data in the BufferChunk. 
		/// Also return the compression data which is the remaining bytes at the end of the byte[].
		/// </summary>
		/// <param name="mt"></param>
		/// <param name="bc"></param>
		/// <param name="compressionData"></param>
		public static void ReconstituteAudioFormat(MediaTypeWaveFormatEx mt, BufferChunk bc, out byte[] compressionData)
		{
			mt.WaveFormatEx.FormatTag = (ushort)NextInt16(bc);      
			mt.WaveFormatEx.Channels = (ushort)NextInt16(bc);       
			mt.WaveFormatEx.SamplesPerSec = (uint)NextInt32(bc);  
			mt.WaveFormatEx.AvgBytesPerSec = (uint)NextInt32(bc); 
			mt.WaveFormatEx.BlockAlign = (ushort)NextInt16(bc);     
			mt.WaveFormatEx.BitsPerSample = (ushort)NextInt16(bc);  
			mt.WaveFormatEx.Size = (ushort)NextInt16(bc);   
			compressionData = new byte[mt.WaveFormatEx.Size];
			for (int i=0;i<mt.WaveFormatEx.Size;i++)
				compressionData[i] = bc.NextByte();
		}

		/// <summary>
		/// Fill in the video-specific parts of the MediaType from the data in the BufferChunk. 
		/// Also return the compression data which is the remaining bytes at the end of the byte[].
		/// </summary>
		/// <param name="mt"></param>
		/// <param name="bc"></param>
		/// <param name="compressionData"></param>
		public static void ReconstituteVideoFormat(MediaTypeVideoInfo mt, BufferChunk bc, out byte[] compressionData)
		{
			VIDEOINFOHEADER vi;
			RECT s;
			s.left = NextInt32(bc);
			s.top = NextInt32(bc);
			s.right = NextInt32(bc);
			s.bottom = NextInt32(bc);
			vi.Source = s;
			RECT t;
			t.left = NextInt32(bc);
			t.top = NextInt32(bc);
			t.right = NextInt32(bc);
			t.bottom = NextInt32(bc);
			vi.Target = t;
			vi.BitRate = (uint)NextInt32(bc);
			vi.BitErrorRate = (uint)NextInt32(bc);
			vi.AvgTimePerFrame = bc.NextUInt64();
			BITMAPINFOHEADER bih;
			bih.Size = (uint)NextInt32(bc);
			bih.Width = NextInt32(bc);
			bih.Height = NextInt32(bc);
			bih.Planes = (ushort)NextInt16(bc);
			bih.BitCount = (ushort)NextInt16(bc);
			bih.Compression = (uint)NextInt32(bc);
			bih.SizeImage = (uint)NextInt32(bc);
			bih.XPelsPerMeter = NextInt32(bc);
			bih.YPelsPerMeter = NextInt32(bc);
			bih.ClrUsed = (uint)NextInt32(bc);
			bih.ClrImportant = (uint)NextInt32(bc);
			vi.BitmapInfo = bih;
			mt.VideoInfo = vi;
			compressionData = new byte[bc.Length];
			for (int i=0; i< compressionData.Length; i++)
			{
				compressionData[i] = bc.NextByte();
			}
		}

		/// <summary>
		/// Take the DirectShow header off the frame and return the remainder.
		/// Also indicate whether the keyframe flag is set.
		/// </summary>
		/// <param name="frame"></param>
		/// <returns></returns>
		public static BufferChunk FrameToSample(BufferChunk frame, out bool keyframe)
		{
			short headerSize = frame.NextInt16();  //first short tells us the header size
            BufferChunk header = null;
            header = frame.NextBufferChunk(headerSize);

			AM_SAMPLE2_PROPERTIES amsp = ReconstituteSampleProperties(header);
			if ((amsp.dwSampleFlags&1)>0)
				keyframe = true;
			else
				keyframe = false;
			//Debug.WriteLine("FrameToSample returning sample of size=" + frame.Length);
			return frame;
		}

        public static void DebugPrintAudioFormat(MediaTypeWaveFormatEx mt)
        {
            Debug.WriteLine("  AvgBytesPerSec=" + mt.WaveFormatEx.AvgBytesPerSec.ToString());
            Debug.WriteLine("  BitsPerSample=" + mt.WaveFormatEx.BitsPerSample.ToString());
            Debug.WriteLine("  BlockAlign=" + mt.WaveFormatEx.BlockAlign.ToString());
            Debug.WriteLine("  Channels=" + mt.WaveFormatEx.Channels.ToString());
            Debug.WriteLine("  FormatTag=" + mt.WaveFormatEx.FormatTag.ToString());
            Debug.WriteLine("  SamplesPerSec=" + mt.WaveFormatEx.SamplesPerSec.ToString());
            Debug.WriteLine("  Size=" + mt.WaveFormatEx.Size.ToString());
        }

        public static void DebugPrintVideoFormat(MediaTypeVideoInfo mt)
        {
            Debug.WriteLine("  AvgTimePerFrame=" + mt.VideoInfo.AvgTimePerFrame.ToString());
            Debug.WriteLine("  BitErrorRate=" + mt.VideoInfo.BitErrorRate.ToString());
            Debug.WriteLine("  BitRate=" + mt.VideoInfo.BitRate.ToString());
            Debug.WriteLine("  Source.top=" + mt.VideoInfo.Source.top.ToString());
            Debug.WriteLine("  Source.left=" + mt.VideoInfo.Source.left.ToString());
            Debug.WriteLine("  Source.bottom=" + mt.VideoInfo.Source.bottom.ToString());
            Debug.WriteLine("  Source.right=" + mt.VideoInfo.Source.right.ToString());
            Debug.WriteLine("  Target.top=" + mt.VideoInfo.Target.top.ToString());
            Debug.WriteLine("  Target.left=" + mt.VideoInfo.Target.left.ToString());
            Debug.WriteLine("  Target.bottom=" + mt.VideoInfo.Target.bottom.ToString());
            Debug.WriteLine("  Target.right=" + mt.VideoInfo.Target.right.ToString());
            Debug.WriteLine("  BitmapInfo.Height=" + mt.VideoInfo.BitmapInfo.Height.ToString());
            Debug.WriteLine("  BitmapInfo.Planes=" + mt.VideoInfo.BitmapInfo.Planes.ToString());
            Debug.WriteLine("  BitmapInfo.BitCount=" + mt.VideoInfo.BitmapInfo.BitCount.ToString());
            Debug.WriteLine("  BitmapInfo.ClrImportant=" + mt.VideoInfo.BitmapInfo.ClrImportant.ToString());
            Debug.WriteLine("  BitmapInfo.ClrUsed=" + mt.VideoInfo.BitmapInfo.ClrUsed.ToString());
            Debug.WriteLine("  BitmapInfo.Compression=" + mt.VideoInfo.BitmapInfo.Compression.ToString());
            Debug.WriteLine("  BitmapInfo.Size=" + mt.VideoInfo.BitmapInfo.Size.ToString());
            Debug.WriteLine("  BitmapInfo.SizeImage=" + mt.VideoInfo.BitmapInfo.SizeImage.ToString());
            Debug.WriteLine("  BitmapInfo.Width=" + mt.VideoInfo.BitmapInfo.Width.ToString());
            Debug.WriteLine("  BitmapInfo.XPelsPerMeter=" + mt.VideoInfo.BitmapInfo.XPelsPerMeter.ToString());
            Debug.WriteLine("  BitmapInfo.YPelsPerMeter=" + mt.VideoInfo.BitmapInfo.YPelsPerMeter.ToString());

        }


        public static void DebugPrintBaseMediaType(MediaType mt)
        {
            Debug.WriteLine("*********" + mt.ToString() + "************");
            Debug.WriteLine("MajorType=" + Enum.GetName(typeof(MajorType), mt.MajorType));
            Debug.WriteLine("SubType=" + Enum.GetName(typeof(SubType), mt.SubType));
            Debug.WriteLine("FormatType=" + Enum.GetName(typeof(FormatType), mt.FormatType));
            Debug.WriteLine("SampleSize=" + mt.SampleSize.ToString());
            Debug.WriteLine("FixedSizeSamples=" + mt.FixedSizeSamples.ToString());
            Debug.WriteLine("TemporalCompression=" + mt.TemporalCompression.ToString());
        }


		#endregion Public Static

		#region Private Static

        /// <summary>
        /// Return a new ProfileData instance containing MediaTypes and codec private data as determined
        /// by the audio and video frames given.  One, but not both, frames may be null.
        /// </summary>
        /// <param name="aframe"></param>
        /// <param name="vframe"></param>
        /// <returns></returns>
        private static ProfileData FramesToProfileData(byte[] aframe, byte[] vframe)
        {
            if ((aframe == null) && (vframe == null))
                return null;

            byte[] audioCompressionData;
            byte[] videoCompressionData;
            MediaTypeWaveFormatEx amt = AudioMediaTypeFromFrame(aframe, out audioCompressionData);
            MediaTypeVideoInfo vmt = VideoMediaTypeFromFrame(vframe, out videoCompressionData);
            return new ProfileData(vmt, videoCompressionData, amt, audioCompressionData);
        }

        private static ProfileData AudioFrameToProfileData(byte[] aframe)
        { 
            if (aframe == null)
                return null;

            byte[] audioCompressionData;
            MediaTypeWaveFormatEx amt = AudioMediaTypeFromFrame(aframe, out audioCompressionData);
            return new ProfileData(amt, audioCompressionData);       
        }


		/// <summary>
		/// Given a byte[] containing an audio frame, return the audio MediaType.
		/// </summary>
		/// <param name="frame"></param>
		/// <param name="compressionData"></param>
		/// <returns></returns>
		private static MediaTypeWaveFormatEx AudioMediaTypeFromFrame(byte[] frame, out byte[] compressionData)
		{
            if (frame == null)
            {
                compressionData = null;
                return null;
            }

			BufferChunk bc = new BufferChunk(frame);
			short headerSize = bc.NextInt16();  //first short tells us the header size
			BufferChunk header = bc.NextBufferChunk(headerSize);
			
			//The header contains a custom serialization of AM_SAMPLE2_PROPERTIES followed by 
			// AM_MEDIA_TYPE and an optional format type.

			//AM_SAMPLE2_PROPERTIES
			BufferChunk AmSample2Properties = header.NextBufferChunk(48); 

			//AM_MEDIA_TYPE 
			MediaTypeWaveFormatEx amt = new MediaTypeWaveFormatEx();
			ReconstituteBaseMediaType((MediaType)amt, header);
			compressionData = null;
			if (amt.FormatType == FormatType.WaveFormatEx)
			{
				ReconstituteAudioFormat(amt, header, out compressionData);
			}
			return amt;
		}


		/// <summary>
		/// Given a BufferChunk containing AM_SAMPLE2_PROPERTIES, return a filled-in struct.
		/// </summary>
		/// <param name="bc"></param>
		/// <returns></returns>
		private static AM_SAMPLE2_PROPERTIES ReconstituteSampleProperties(BufferChunk bc)
		{
			AM_SAMPLE2_PROPERTIES asp;
			asp.cbData = (uint)NextInt32(bc);
			asp.dwTypeSpecificFlags = (uint)NextInt32(bc);
			asp.dwSampleFlags = (uint)NextInt32(bc);
			asp.lActual = NextInt32(bc);
			asp.tStart = bc.NextUInt64();
			asp.tStop = bc.NextUInt64();
			asp.dwStreamId = (uint)NextInt32(bc);
			asp.pMediaType = (IntPtr)NextInt32(bc);
			asp.pbBuffer = (IntPtr)NextInt32(bc);
			asp.cbBuffer = NextInt32(bc);
			return asp;
		}

		/// <summary>
		/// Given a byte[] containing a video frame, return the video MediaType.
		/// </summary>
		/// <param name="frame"></param>
		/// <param name="compressionData"></param>
		/// <returns></returns>
		private static MediaTypeVideoInfo VideoMediaTypeFromFrame(byte[] frame, out byte[] compressionData)
		{
            if (frame == null)
            {
                compressionData = null;
                return null;
            }

			BufferChunk bc = new BufferChunk(frame);
			short headerSize = bc.NextInt16();  //first short tells us the header size
			BufferChunk header = bc.NextBufferChunk(headerSize);
			
			//The header contains a custom serialization of AM_SAMPLE2_PROPERTIES followed by 
			// AM_MEDIA_TYPE and an optional format type.

			//AM_SAMPLE2_PROPERTIES
			BufferChunk AmSample2Properties = header.NextBufferChunk(48); 

			//AM_MEDIA_TYPE 
			MediaTypeVideoInfo vmt = new MediaTypeVideoInfo();
			ReconstituteBaseMediaType((MediaType)vmt, header);
			compressionData = null;
			if (vmt.FormatType == FormatType.VideoInfo)
			{
				ReconstituteVideoFormat(vmt, header, out compressionData);
			}
			return vmt;
		}

		/// <summary>
		/// Extract a Guid from the next bytes in a BufferChunk
		/// </summary>
		/// <param name="bc"></param>
		/// <returns></returns>
		private static Guid NextGuid(BufferChunk bc)
		{
			int a = NextInt32(bc);
			short b = NextInt16(bc);
			short c = NextInt16(bc);
			byte[] d = (byte[])bc.NextBufferChunk(8);
			return new Guid(a,b,c,d);
		}

		/// <summary>
		/// Extract a Int32 from the next bytes in a BufferChunk.
		/// </summary>
		/// <param name="bc"></param>
		/// <returns></returns>
		private static Int32 NextInt32(BufferChunk bc)
		{
			Int32 i = (Int32)bc.NextByte();
			i += (Int32)(bc.NextByte() << 8);
			i += (Int32)(bc.NextByte() << 16);
			i += (Int32)(bc.NextByte() << 24);
			return i;
		}

		/// <summary>
		/// Extract a Int16 from the next bytes in a BufferChunk.
		/// </summary>
		/// <param name="bc"></param>
		/// <returns></returns>
		private static Int16 NextInt16(BufferChunk bc)
		{
			Int16 i = (Int16)bc.NextByte();
			i += (Int16)(bc.NextByte() << 8);
			return i;
		}

		#endregion Private Static

		#region AM_SAMPLE2_PROPERTIES Struct

		/// <summary>
		/// This is a struct we need from DirectShow..
		/// </summary>
		private struct AM_SAMPLE2_PROPERTIES 
		{
			public UInt32	cbData;
			public UInt32	dwTypeSpecificFlags;
			public UInt32	dwSampleFlags;
			public Int32	lActual;
			public UInt64	tStart;
			public UInt64	tStop;
			public UInt32	dwStreamId;
			public IntPtr	pMediaType;
			public IntPtr	pbBuffer;
			public Int32	cbBuffer;
		}

		#endregion AM_SAMPLE2_PROPERTIES Struct

        #region Test Code

        ///// <summary>
        ///// Require input segment contains exactly one audio and one video, and the MediaType 
        ///// of the streams do not change during the segment.  Return the native Windows Media profile
        ///// string, or null for any error.
        ///// </summary>
        ///// <param name="segment"></param>
        ///// <returns></returns>
        //public static String MakeNativeProfile(ArchiveTranscoderJobSegment segment)
        //{
        //    DateTime startDt = DateTime.Parse(segment.StartTime);
        //    DateTime endDt = DateTime.Parse(segment.EndTime);

        //    byte[] aframe = DatabaseUtility.GetFirstFrame(PayloadType.dynamicAudio,segment.AudioDescriptor[0].AudioCname,
        //        segment.AudioDescriptor[0].AudioName,startDt.Ticks,endDt.Ticks);
        //    byte[] vframe = DatabaseUtility.GetFirstFrame(PayloadType.dynamicVideo,segment.VideoDescriptor.VideoCname,
        //        segment.VideoDescriptor.VideoName,startDt.Ticks,endDt.Ticks);
        //    if ((aframe==null) || (vframe==null))
        //    {
        //        return null;
        //    }

        //    return FramesToProfile(aframe,vframe);
        //}

        ///// <summary>
        ///// Given a Stream ID for a video stream, return the native Windows Media Profile string, or null.
        ///// </summary>
        ///// <param name="streamID"></param>
        ///// <returns></returns>
        //public static String MakeNativeVideoProfile(int streamID)
        //{
        //    byte[] vframe = DatabaseUtility.GetFirstFrame(streamID);
        //    if (vframe==null)
        //        return null;

        //    return VideoFrameToProfile(vframe);
        //}

        ///// <summary>
        ///// Given a Stream ID for an audio stream, return the native Windows Media profile string, or null.
        ///// </summary>
        ///// <param name="streamID"></param>
        ///// <returns></returns>
        //public static String MakeNativeAudioProfile(int streamID)
        //{
        //    byte[] aframe = DatabaseUtility.GetFirstFrame(streamID);
        //    if (aframe==null)
        //        return null;

        //    return AudioFrameToProfile(aframe);
        //}

        ///// <summary>
        ///// Return the AM_MEDIA_TYPE portion as a byte[] from the first frame for the given streamID
        ///// </summary>
        ///// <param name="streamID"></param>
        ///// <returns></returns>
        //public static byte[] GetRawMediaType(int streamID)
        //{
        //    byte[] vframe = DatabaseUtility.GetFirstFrame(streamID);
        //    BufferChunk bc = new BufferChunk(vframe);
        //    short headerSize = bc.NextInt16();  //first short tells us the header size
        //    //The header contains a custom serialization of AM_SAMPLE2_PROPERTIES followed by 
        //    // AM_MEDIA_TYPE and an optional format type.	
        //    //AM_SAMPLE2_PROPERTIES is always 48 bytes.
        //    headerSize -= 48;
        //    short index = 50; //2+48

        //    byte[] header = new byte[headerSize];
        //    Array.Copy(bc.Buffer,index,header,0,headerSize);
        //    return header;
        //}

        ///// <summary>
        ///// Given a byte[] containing a video frame, build a matching Windows Media
        ///// profile string containing a single video stream.
        ///// </summary>
        ///// <param name="vframe"></param>
        ///// <returns></returns>
        //public static String VideoMediaTypeToProfile(MediaTypeVideoInfo vmt, byte[] codecData)
        //{
        //    StringBuilder sb = new StringBuilder();
        //    sb.Append ("<profile version=\"589824\" storageformat=\"1\" name=\"CXP\" description=\"\"> ");
        //    BuildVideoStreamConfig(sb, vmt, codecData,1);
        //    sb.Append ("</profile>");
        //    return sb.ToString();
        //}

        ///// <summary>
        ///// Given a byte[] containing a video frame, build a matching Windows Media
        ///// profile string containing a single video stream.
        ///// </summary>
        ///// <param name="vframe"></param>
        ///// <returns></returns>
        //private static String VideoFrameToProfile(byte[] vframe)
        //{
        //    byte[] videoCompressionData;
        //    MediaTypeVideoInfo vmt = VideoMediaTypeFromFrame(vframe, out videoCompressionData);
        //    StringBuilder sb = new StringBuilder();
        //    sb.Append ("<profile version=\"589824\" storageformat=\"1\" name=\"CXP\" description=\"\"> ");
        //    BuildVideoStreamConfig(sb, vmt, videoCompressionData,1);
        //    sb.Append ("</profile>");
        //    //Debug.WriteLine(sb.ToString());
        //    return sb.ToString();
        //}

        ///// <summary>
        ///// Given a byte[] containing an audio frame, build a matching Windows Media
        ///// profile string containing a single audio stream.
        ///// </summary>
        ///// <param name="aframe"></param>
        ///// <returns></returns>
        //private static String AudioFrameToProfile(byte[] aframe)
        //{
        //    byte[] audioCompressionData;
        //    MediaTypeWaveFormatEx amt = AudioMediaTypeFromFrame(aframe, out audioCompressionData);
        //    StringBuilder sb = new StringBuilder();
        //    sb.Append ("<profile version=\"589824\" storageformat=\"1\" name=\"CXP\" description=\"\"> ");
        //    BuildAudioStreamConfig(sb, amt, audioCompressionData,1);
        //    sb.Append ("</profile>");
        //    //Debug.WriteLine(sb.ToString());
        //    return sb.ToString();
        //}

        ///// <summary>
        ///// Given byte[]s containing audio and video frames, build the matching Windows Media
        ///// profile string containing both audio and video streams.
        ///// </summary>
        ///// <param name="aframe"></param>
        ///// <param name="vframe"></param>
        ///// <returns></returns>
        //private static String FramesToProfile(byte[] aframe, byte[] vframe)
        //{
        //    byte[] audioCompressionData;
        //    byte[] videoCompressionData;
        //    MediaTypeWaveFormatEx amt = AudioMediaTypeFromFrame(aframe, out audioCompressionData);
        //    MediaTypeVideoInfo vmt = VideoMediaTypeFromFrame(vframe, out videoCompressionData);
        //    return MediaTypeToProfile(amt, vmt, audioCompressionData, videoCompressionData);
        //}


        ///// <summary>
        ///// Given Audio and Video MediaTypes, return the matching Windows Media profile
        ///// string contining both audio and video streams.
        ///// </summary>
        ///// <param name="amt"></param>
        ///// <param name="vmt"></param>
        ///// <param name="aData"></param>
        ///// <param name="vData"></param>
        ///// <returns></returns>
        //private static String MediaTypeToProfile(MediaTypeWaveFormatEx amt, MediaTypeVideoInfo vmt, byte[] aData, byte[] vData)
        //{
        //    StringBuilder sb = new StringBuilder();
        //    sb.Append ("<profile version=\"589824\" storageformat=\"1\" name=\"CXP\" description=\"\"> ");
        //    BuildAudioStreamConfig(sb, amt, aData,1);
        //    BuildVideoStreamConfig(sb, vmt, vData,2);
        //    sb.Append ("</profile>");
        //    //Debug.WriteLine(sb.ToString());
        //    return sb.ToString();
        //}

        ///// <summary>
        ///// Given the Audio MediaType, add the audio 'StreamConfig' section of the matching Windows Media profile
        ///// to the StringBuilder.
        ///// </summary>
        ///// <param name="sb"></param>
        ///// <param name="amt"></param>
        ///// <param name="data"></param>
        ///// <param name="streamNumber"></param>
        //private static void BuildAudioStreamConfig(StringBuilder sb, MediaTypeWaveFormatEx amt, byte[] data, int streamNumber)
        //{
        //    String fss = amt.FixedSizeSamples==true?"1":"0";
        //    String tc = amt.TemporalCompression==true?"1":"0";
        //    uint bitrate = amt.WaveFormatEx.AvgBytesPerSec * 8; //changed from a hardcoded value of 20050.
        //    sb.Append("<streamconfig majortype=\"{" + amt.mt.majortype.ToString() + "}\" ");
        //    sb.Append("streamnumber=\"" + streamNumber.ToString() + "\" streamname=\"Audio Stream\" inputname=\"Audio409\" " +
        //        "bitrate=\"" + bitrate.ToString() + "\" bufferwindow=\"5000\" reliabletransport=\"0\" decodercomplexity=\"\" " +
        //        "rfc1766langid=\"en-us\" > ");
        //    sb.Append("<wmmediatype subtype=\"{" + amt.mt.subtype.ToString() + "}\" " +
        //        "bfixedsizesamples=\"" + fss + "\" " +
        //        "btemporalcompression=\"" + tc + "\" " +
        //        "lsamplesize=\"" + amt.SampleSize.ToString() +"\"> "); 
        //    sb.Append("<waveformatex wFormatTag=\"" + amt.WaveFormatEx.FormatTag.ToString() + "\" " +
        //        "nChannels=\"" + amt.WaveFormatEx.Channels.ToString() + "\" " + 
        //        "nSamplesPerSec=\"" + amt.WaveFormatEx.SamplesPerSec.ToString() + "\" " + 
        //        "nAvgBytesPerSec=\"" + amt.WaveFormatEx.AvgBytesPerSec.ToString() + "\" " + 
        //        "nBlockAlign=\"" + amt.WaveFormatEx.BlockAlign.ToString() + "\" " + 
        //        "wBitsPerSample=\"" + amt.WaveFormatEx.BitsPerSample.ToString() + "\" " + 
        //        "codecdata=\"" + Utility.ByteArrayToHexString(data) + "\"/> " );
        //    sb.Append("</wmmediatype> </streamconfig> ");
        //}

        ///// <summary>
        ///// Given the video MediaType, add the video 'StreamConfig' section of the matching Windows Media profile
        ///// to the StringBuilder.
        ///// </summary>
        ///// <param name="sb"></param>
        ///// <param name="vmt"></param>
        ///// <param name="data"></param>
        ///// <param name="streamNumber"></param>
        //private static void BuildVideoStreamConfig(StringBuilder sb, MediaTypeVideoInfo vmt, byte[] data, int streamNumber)
        //{
        //    String fss = vmt.FixedSizeSamples==true?"1":"0";
        //    String tc = vmt.TemporalCompression==true?"1":"0";

        //    sb.Append("<streamconfig majortype=\"{" + vmt.mt.majortype.ToString() + "}\" ");
        //    sb.Append("streamnumber=\"" + streamNumber.ToString() + "\" streamname=\"Video Stream\" inputname=\"Video409\" " +
        //        "bitrate=\"" + vmt.VideoInfo.BitRate.ToString() + "\" bufferwindow=\"5000\" " +
        //        "reliabletransport=\"0\" decodercomplexity=\"AU\" " +
        //        "rfc1766langid=\"en-us\" > ");

        //    sb.Append("<videomediaprops maxkeyframespacing=\"80000000\" " +
        //        "quality=\"90\"/> ");
        //    sb.Append("<wmmediatype subtype=\"{" + vmt.mt.subtype.ToString() + "}\" " +
        //        "bfixedsizesamples=\"" + fss + "\" " +
        //        "btemporalcompression=\"" + tc + "\" " +
        //        "lsamplesize=\"" + vmt.SampleSize.ToString() + "\"> ");
        //    sb.Append("<videoinfoheader dwbitrate=\"" + vmt.VideoInfo.BitRate.ToString() + "\" " +
        //        "dwbiterrorrate=\"" + vmt.VideoInfo.BitErrorRate.ToString() + "\" " +
        //        "avgtimeperframe=\"" + vmt.VideoInfo.AvgTimePerFrame.ToString() + "\"> " +
        //        "<rcsource left=\"" + vmt.VideoInfo.Source.left.ToString() + "\" " +
        //        "top=\"" + vmt.VideoInfo.Source.top.ToString() + "\" " +
        //        "right=\"" + vmt.VideoInfo.Source.right.ToString() + "\" " +
        //        "bottom=\"" + vmt.VideoInfo.Source.bottom.ToString() + "\"/> " +
        //        "<rctarget left=\"" + vmt.VideoInfo.Target.left.ToString() + "\" " +
        //        "top=\"" + vmt.VideoInfo.Target.top.ToString() + "\" " +
        //        "right=\"" + vmt.VideoInfo.Target.right.ToString() + "\" " +
        //        "bottom=\"" + vmt.VideoInfo.Target.bottom.ToString() + "\"/> ");
        //    sb.Append("<bitmapinfoheader biwidth=\"" + vmt.VideoInfo.BitmapInfo.Width.ToString() + "\" " +
        //        "biheight=\"" + vmt.VideoInfo.BitmapInfo.Height.ToString() + "\" " +
        //        "biplanes=\"" + vmt.VideoInfo.BitmapInfo.Planes.ToString() + "\" " +
        //        "bibitcount=\"" + vmt.VideoInfo.BitmapInfo.BitCount.ToString() + "\" " +
        //        "bicompression=\"" + vmt.SubType.ToString() + "\" " +
        //        "bisizeimage=\"" + vmt.VideoInfo.BitmapInfo.SizeImage.ToString() + "\" " +
        //        "bixpelspermeter=\"" + vmt.VideoInfo.BitmapInfo.XPelsPerMeter.ToString() + "\" " +
        //        "biypelspermeter=\"" + vmt.VideoInfo.BitmapInfo.YPelsPerMeter.ToString() + "\" " +
        //        "biclrused=\"" + vmt.VideoInfo.BitmapInfo.ClrUsed.ToString() + "\" " +
        //        "biclrimportant=\"" + vmt.VideoInfo.BitmapInfo.ClrImportant.ToString() + "\" " +
        //        "codecdata=\"" + Utility.ByteArrayToHexString(data) + "\" " +
        //        "/> " +
        //        "</videoinfoheader> ");
        //    sb.Append("</wmmediatype> </streamconfig> ");

        //}
        #endregion Test Code
    }

    #region ProfileData Class

    /// <summary>
    /// Encapsulate MediaType and compression data used to build profiles.
    /// </summary>
    public class ProfileData
    {
        private MediaTypeVideoInfo videoMediaType;
        private byte[] videoCodecData;
        private MediaTypeWaveFormatEx audioMediaType;
        private byte[] audioCodecData;

        private Guid videoCodecGuid;
        private int height;
        private int width;
        private uint bitrate;
        private uint bufferwindow;
        private ulong avgtimeperframe;

        public MediaTypeVideoInfo VideoMediaType
        {
            get { return videoMediaType; }
        }

        public byte[] VideoCodecData
        {
            get { return videoCodecData; }
        }

        public MediaTypeWaveFormatEx AudioMediaType
        {
            get { return audioMediaType; }
        }

        public byte[] AudioCodecData
        {
            get { return audioCodecData; }
        }

        public Guid VideoCodecGuid
        {
            get { return videoCodecGuid; }
        }

        public int Height
        {
            get { return height; }
        }

        public int Width
        {
            get { return width; }
        }

        public uint BitRate
        {
            get { return bitrate; }
        }

        public uint BufferWindow
        {
            get { return bufferwindow; }
        }

        public ulong AvgTimePerFrame
        {
            get { return avgtimeperframe; }
        }

        public ProfileData(MediaTypeVideoInfo videoMediaType, byte[] videoCodecData)
        {
            init();
            this.videoMediaType = videoMediaType;
            this.videoCodecData = videoCodecData;
        }

        public ProfileData(MediaTypeWaveFormatEx audioMediaType, byte[] audioCodecData)
        {
            init();
            this.audioMediaType = audioMediaType;
            this.audioCodecData = audioCodecData;
        }

        public ProfileData(MediaTypeVideoInfo videoMediaType, byte[] videoCodecData, MediaTypeWaveFormatEx audioMediaType, byte[] audioCodecData)
        {
            init();
            this.videoMediaType = videoMediaType;
            this.videoCodecData = videoCodecData;
            this.audioMediaType = audioMediaType;
            this.audioCodecData = audioCodecData;
        }

        /// <summary>
        /// If the video mediatype is null, this data will determine the preferred video codec and settings.
        /// </summary>
        /// <param name="codecId"></param>
        /// <param name="h"></param>
        /// <param name="w"></param>
        /// <param name="bitrate"></param>
        /// <param name="bufferwindow"></param>
        public void SetSecondaryVideoCodecData(Guid codecId, int h, int w, uint bitrate, uint bufferwindow, ulong avgtimeperframe)
        {
            this.videoCodecGuid = codecId;
            this.height = h;
            this.width = w;
            this.bitrate = bitrate;
            this.bufferwindow = bufferwindow;
            this.avgtimeperframe = avgtimeperframe;
        }

        private void init()
        {
            this.audioMediaType = null;
            this.audioCodecData = null;
            this.videoMediaType = null;
            this.videoCodecData = null;
            this.videoCodecGuid = Guid.Empty;
            this.height = 0;
            this.width = 0;
            this.bitrate = 0;
            this.bufferwindow = 0;
        }
    }

    #endregion ProfileData Class

}
