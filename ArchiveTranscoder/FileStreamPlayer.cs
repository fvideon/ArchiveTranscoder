using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;
using MSR.LST;
using MSR.LST.MDShow;
using UW.CSE.ManagedWM;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Given a WMV or WMA file, maintain a IWMSyncReader on the file, returning 
	/// compressed or uncompressed samples and timestamps, reading from the beginning of the file to the
	/// end, or until the timestamp exceeds the end specified to the constructor.
	/// Timestamps returned are absolute.  The file must contain only a single stream.
	/// </summary>
	public class FileStreamPlayer
	{
		#region Members

		private long start;
		private long end;
		private String filename;
		private IWMSyncReader reader;
		private MediaTypeVideoInfo vmt = null;
		private MediaTypeWaveFormatEx amt = null;
		private bool outOfData;
		private ulong duration;
		private Guid guid;
		private int streamID;

		#endregion Members

		#region Properties

		public int StreamID
		{
			get { return streamID; }
		}

		public MediaTypeVideoInfo VideoMediaType
		{
			get {return vmt;}
		}

		public MediaTypeWaveFormatEx AudioMediaType
		{
			get {return amt;}
		}

		public ulong Duration
		{
			get { return duration; }
		}

		public long StartTime
		{
			get { return start; }
		}

		public Guid xGuid
		{
			get { return guid; }
		}

		#endregion Properties

		#region Ctor/Dtor

		public FileStreamPlayer(String filename, long start, long end, bool compressed, int streamID)
		{
			this.streamID = streamID;
			this.filename = filename;
			this.start = start;
			this.end = end;
			this.duration = (ulong)(end-start);
			outOfData = false;
			this.guid = Guid.NewGuid();

			//create IWMSyncReader and open the file.
			uint hr = WMFSDKFunctions.WMCreateSyncReader(null,0,out reader);
			IntPtr fn = Marshal.StringToCoTaskMemUni(filename);
			reader.Open(fn);
			Marshal.FreeCoTaskMem(fn);

			//Verify that the file contains one stream.
			uint outputcnt;
			reader.GetOutputCount(out outputcnt);
			Debug.Assert(outputcnt==1);

			//Extract the MediaType for the stream.
			uint cmt=0;
			IntPtr ipmt;
			IWMOutputMediaProps outputProps;
			reader.GetOutputProps(0,out outputProps);
			outputProps.GetMediaType(IntPtr.Zero,ref cmt);
			ipmt = Marshal.AllocCoTaskMem((int)cmt);
			outputProps.GetMediaType(ipmt,ref cmt);
			byte[] bmt = new byte[cmt];
			Marshal.Copy(ipmt,bmt,0,(int)cmt);
			BufferChunk bc = new BufferChunk(bmt);
			byte[] cd;

			GUID majorTypeGUID;
			outputProps.GetType(out majorTypeGUID);
			if (WMGuids.ToGuid(majorTypeGUID)==WMGuids.WMMEDIATYPE_Video)
			{
				vmt = new MediaTypeVideoInfo();
				ProfileUtility.ReconstituteBaseMediaType((MediaType)vmt,bc);
				ProfileUtility.ReconstituteVideoFormat(vmt,bc,out cd);
                //Note: This is a special case which we would like to generalize:  The default output format for the 
                //12bpp video was found not to return any uncompressed samples.  Setting this particular case to RGB 24 fixed it.
                if ((!compressed) && (vmt.VideoInfo.BitmapInfo.BitCount==12)) {
                    SetVideoOutputProps();
                }
			}
			else if (WMGuids.ToGuid(majorTypeGUID)==WMGuids.WMMEDIATYPE_Audio)
			{
				amt = new MediaTypeWaveFormatEx();
				ProfileUtility.ReconstituteBaseMediaType((MediaType)amt,bc);
				ProfileUtility.ReconstituteAudioFormat(amt,bc,out cd);
			}

			//if compressed is set, retrieve stream samples
			if (compressed)
			{
				reader.SetReadStreamSamples(1,1);
			}

		}

        /// <summary>
        /// It was found that sometimes the default output props returns zero samples.
        /// We whould like to query the reader to find out which formats are supported.
        /// </summary>
        private void SetVideoOutputProps() {
            //Enumerate supported output formats
            uint formatcount;
            reader.GetOutputFormatCount(0, out formatcount);
            for (uint j = 0; j < formatcount; j++) {
                IWMOutputMediaProps props;
                reader.GetOutputFormat(0, j, out props);
                uint cmt = 0;
                props.GetMediaType(IntPtr.Zero, ref cmt);
                IntPtr ipmt = Marshal.AllocCoTaskMem((int)cmt);
                props.GetMediaType(ipmt, ref cmt);
                byte[] bmt = new byte[cmt];
                MediaTypeVideoInfoAlt mt = (MediaTypeVideoInfoAlt)Marshal.PtrToStructure(ipmt, typeof(MediaTypeVideoInfoAlt));
                string fourcc = new string(mt.compression);
                //Set the format to RGB24 if available:
                if ((fourcc == "\0\0\0\0") &&
                    (mt.bitCount == 24)) {
                    Debug.WriteLine("Setting output format to RGB 24");
                    reader.SetOutputProps(0, props);
                    break;
                }
                Marshal.FreeCoTaskMem(ipmt);
            }
        }

		~FileStreamPlayer()
		{
			if (File.Exists(filename))
			{
				File.Delete(filename);
			}
		}

		#endregion Ctor/Dtor

		#region Public Method


        public bool GetNextSample(out BufferChunk sample, out long timestamp)
        {
            bool junk;
            return GetNextSample(out sample, out timestamp, out junk);
        }

		/// <summary>
		/// Return the next sample from the stream.  Return false if there are no more samples.
		/// </summary>
		/// <param name="sample"></param>
		/// <param name="timestamp"></param>
		/// <returns></returns>
		public bool GetNextSample(out BufferChunk sample, out long timestamp, out bool keyframe)
		{
			sample = null;
			timestamp = 0;
            keyframe = false;
			
			if (outOfData)
				return false;

			ulong stime;
			ulong sduration;
			uint flags;
			uint outnum;
			ushort streamnum;
			INSSBuffer buf;
			try
			{
				reader.GetNextSample(0,out buf,out stime,out sduration,out flags,out outnum,out streamnum);
			}
			catch (Exception e)
			{
                //0xC00D0BCF is 'no more samples'
                if (e.Message != "Exception from HRESULT: 0xC00D0BCF")
                    Debug.WriteLine("Exception while reading samples: " + e.ToString());

				//This exception seems to be the only way to tell when a stream is exhausted?
				//Debug.WriteLine("Exception in GetNextSample while reading: " + this.filename + " " + e.ToString());
				outOfData = true;
				return false;
			}

			if ((long)stime + start > end)
			{
				return false;
			}
			uint buflen;
			IntPtr ibuf;
			buf.GetLength(out buflen);
			buf.GetBuffer(out ibuf);
			byte[] outbuf = new byte[buflen];
			Marshal.Copy(ibuf,outbuf,0,(int)buflen);
			sample = new BufferChunk(outbuf);
			timestamp = (long)stime + start;

            keyframe = false;
            if ((flags & WMGuids.WM_SF_CLEANPOINT) != 0)
                keyframe = true;

            //Debug.WriteLine("FileStreamPlayer.GetNextSample. flags=" + flags.ToString());

			return true;
		}

		#endregion Public Method

        /// <summary>
        /// An alternate VideoInfo. It is a bit simpler to deal with this in some cases.
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct MediaTypeVideoInfoAlt {
            public GUID majorType;
            public GUID subType;
            public int fixedSizeSamples;
            public int temporalCompression;
            public int sampleSize;
            public GUID formatType;
            public int punk;
            public int cbFormat;
            public int pbFormat;
            //VideoInfoheader
            public int sourceLeft;
            public int sourceTop;
            public int sourceRight;
            public int sourceBottom;
            public int targetLeft;
            public int targetTop;
            public int targetRight;
            public int targetBottom;
            public uint bitRate;
            public uint bitErrorRate;
            public ulong avgTimePerFrame;
            //bitmapinfoheader
            public uint size;
            public int width;
            public int height;
            public short planes;
            public ushort bitCount;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
            public char[] compression; //marshal a fourcc;
            public uint sizeImage;
            public int xPelsPerMeter;
            public int yPelsPerMeter;
            public int clrUsed;
            public uint clrImportant;
            //compression data may follow
        }

	}
}
