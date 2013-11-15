using System;
using UW.CSE.ManagedWM;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.IO;
using MSR.LST.MDShow;
using MSR.LST.Net.Rtp;
using MSR.LST;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Wrap essential Windows Media Format SDK functionality, and handle interop.
	/// </summary>
	public class WMWriter
	{
		#region Properties

		/// <summary>
		/// The WM Writer failed since the last time we checked.
		/// </summary>
		public bool WriteFailed
		{
			get 
			{
				if (writeFailed)
				{	
					writeFailed = false;
					return true;
				}
				else
					return false;
			}
		}
		#endregion

		#region Members

		// WMFSDK class instances:
		private IWMWriterAdvanced		writerAdvanced;
		private IWMWriter				writer;
		private IWMWriterFileSink 		fileSink;		
		private IWMInputMediaProps      audioProps;
		private IWMInputMediaProps      videoProps;
		private IWMProfileManager		profileManager;

		private uint	                audioInput;			//IWMWriter identifies inputs by number .. audio
		private uint	                videoInput;			// .. and video
		private ushort					videoStreamNum;		//Stream numbers are used for writing compressed samples .. video
		private ushort					audioStreamNum;		//	.. audio
		private ushort					scriptStreamNumber;	//	.. presentation script (unused for now)
		private bool					writeFailed;		//Flags that a write excepted since last checked.
		private	uint					scriptBitrate;		//Unused
		
		#endregion Members

		#region Constructor

		public WMWriter()
		{
			writerAdvanced = null;
			writer = null;
			fileSink = null;
			audioProps = null;
			videoProps = null;
			profileManager = null;

			audioInput = 0;
			videoInput = 0;
			videoStreamNum = 0;
			audioStreamNum = 0;
			scriptStreamNumber = 0;
			writeFailed = false;
			scriptBitrate = 0;
		}

		#endregion

		#region Public Methods

		/// <summary>
		/// Create the basic WM writer objects.
		/// </summary>
		/// <returns>true for success</returns>
		public bool Init()
		{
			try
			{
				uint hr = WMFSDKFunctions.WMCreateWriter(null, out writer);
				writerAdvanced = (IWMWriterAdvanced)writer;
				writerAdvanced.SetLiveSource(0);  //false is the default.
			}
			catch (Exception e)
			{
				Debug.WriteLine("Failed to create IWMWriter: " + e.ToString());
				return false;
			}
			return true;
		}

		/// <summary>
		/// Release all the WMFSDK object refs
		/// </summary>
		public void Cleanup()
		{
			writerAdvanced = null;
			writer = null;
			fileSink = null;
			audioProps = null;
			videoProps = null;
			profileManager = null;
		}

		/// <summary>
		/// Prepare the File writer.
		/// </summary>
		/// <param name="filename"></param>
		/// <returns></returns>
		/// Note IWMWriter::SetOutputFilename does about the same thing.
		public bool ConfigFile(String filename)
		{
			if (writerAdvanced == null)
			{
				Debug.WriteLine("WriterAdvanced must exist before ConfigFile is called.");
				return false;
			}

			try
			{
				uint hr = WMFSDKFunctions.WMCreateWriterFileSink(out fileSink);
				IntPtr fn = Marshal.StringToCoTaskMemUni(filename);
				fileSink.Open(fn);
				writerAdvanced.AddSink(fileSink);
				Marshal.FreeCoTaskMem(fn);
			}
			catch (Exception e)
			{
				Debug.WriteLine("Failed to configure FileSink" + e.ToString());
				return false;
			}
			return true;
		}
		
		/// <summary>
		/// Load a WM Profile (system or custom).  Use the index only if prxString and prxFile are empty strings or null.
		/// Use prxFile only if prxString is empty or null.  Note: use of prxString is obsolete.
		/// </summary>
		/// <param name="prxFile"></param>
		/// <param name="prIndex"></param>
		/// <returns></returns>
		public bool ConfigProfile(String prxString, String prxFile, uint prIndex)
		{
			IWMProfile	profile;
			uint hr = WMFSDKFunctions.WMCreateProfileManager(out profileManager);

			if (((prxString==null) || (prxString=="")) && ((prxFile==null) || (prxFile == "")))
			{
				//use system profile
				Guid prg = ProfileIndexToGuid(prIndex);
				if (prg == Guid.Empty)
				{
					profile = null;
					Debug.WriteLine("Unsupported Profile index.");
					return false;
				}
				
				try
				{
					GUID prG = WMGuids.ToGUID(prg);
					profileManager.LoadProfileByID(ref prG,out profile);
				}
				catch (Exception e)
				{
					Debug.WriteLine("Failed to load system profile: " +e.ToString());
					profile = null;
					return false;
				}
			}
			else
			{
				//use custom profile
				profile = LoadCustomProfile(prxString,prxFile);
				if (profile == null)
					return false;
			}
		
			/// Tell the writer to use this profile.
			try
			{
				writer.SetProfile(profile);
				string name = GetProfileName(profile);
				//Debug.WriteLine("Using profile: " + name);
                //this.DebugPrintProfile(profile);
			}
			catch (Exception e)
			{
				Debug.WriteLine("Failed to set writer profile:" + e.ToString());
				profile = null;
				return false;
			}

			/// A slightly confusing point:  Streams are subobjects of the profile, 
			/// while inputs are subobjects of the Writer.  The difference is in the
			/// multi-bitrate scenario where there may be multiple streams per input.
			/// Stream numbers start with 1, while input numbers and stream indices begin at 0.
			/// If we have a profile that supports scripts, we need the stream number of
			/// the script stream.  For audio and video, we just need input number unless we
			/// are going to use the WriteStreamSample method.
			scriptBitrate = 0;
			audioInput = videoInput = 0;
			scriptStreamNumber = 0;
			audioProps = videoProps = null;

			/// stream number for each stream and script bitrate if any.
			uint cStreams;
			IWMStreamConfig streamConfig;
			GUID streamType;
			profile.GetStreamCount(out cStreams);
			for (uint i=0;i<cStreams;i++)
			{
				profile.GetStream(i,out streamConfig);
				streamConfig.GetStreamType(out streamType);
				if (WMGuids.ToGuid(streamType) == WMGuids.WMMEDIATYPE_Script)
				{
					streamConfig.GetStreamNumber(out scriptStreamNumber);
					streamConfig.GetBitrate(out scriptBitrate);
				}
				else if (WMGuids.ToGuid(streamType) == WMGuids.WMMEDIATYPE_Audio)
				{
					streamConfig.GetStreamNumber(out audioStreamNum);
				}
				else if (WMGuids.ToGuid(streamType) == WMGuids.WMMEDIATYPE_Video)
				{
					streamConfig.GetStreamNumber(out videoStreamNum);
				}

			}
			
			return true;
		}

        public bool ConfigProfile(ProfileData profileData)
        {
            return ConfigProfile(profileData, false);
        }

        /// <summary>
        /// Configure the writer profile from the contents of profileData.  Normally we would use the 
        /// MediaType and codec private data to configure new streams, but if that data is unavailable
        /// we may use the codec identifier and secondary encoding parameters.
        /// Use this to configure the profile prior to writing stream samples (compressed samples).
        /// </summary>
        /// <param name="profileData"></param>
        /// <returns>true if all is good, false otherwise</returns>
        public bool ConfigProfile(ProfileData profileData, bool ignoreAudio)
        {
            if (profileData == null)
                return false;

			//Create a IWMProfileManager
			IWMProfileManager pm;
			uint hr = WMFSDKFunctions.WMCreateProfileManager(out pm);

			//create a empty profile
			IWMProfile profile;
			pm.CreateEmptyProfile(WMT_VERSION.WMT_VER_9_0, out profile);

            ushort nextStreamNum = 1;
            if ((profileData.AudioMediaType != null) && (!ignoreAudio))
            {
                MediaTypeWaveFormatEx amt = profileData.AudioMediaType;

                //create a empty stream
                IWMStreamConfig newStreamConfig;
                GUID g = WMGuids.ToGUID(WMGuids.WMMEDIATYPE_Audio);
                profile.CreateNewStream(ref g, out newStreamConfig);

                //set the mediatype
                if (!SetStreamConfigMediaType(newStreamConfig, amt, profileData.AudioCodecData, PayloadType.dynamicAudio))
                {
                    return false;
                }

                //set bitrate, bufferwindow ...
                newStreamConfig.SetBitrate(amt.WaveFormatEx.AvgBytesPerSec * 8);
                newStreamConfig.SetBufferWindow(5000);

                //Add new streamconfig to profile
                try
                {
                    profile.AddStream(newStreamConfig);
                }
                catch { return false; }

                //Keep track of the stream number for writing stream samples
                audioStreamNum = nextStreamNum;
                nextStreamNum++;
            }

            if (profileData.VideoMediaType != null)
            {
                MediaTypeVideoInfo vmt = profileData.VideoMediaType;

                //create a empty stream
                IWMStreamConfig newStreamConfig;
                GUID g = WMGuids.ToGUID(WMGuids.WMMEDIATYPE_Video);
                profile.CreateNewStream(ref g, out newStreamConfig);

                //"fix" vmt so that it will work when the screen codec is used.
                FixMTForScreenCodec(vmt);

                //Set the media type
                if (!SetStreamConfigMediaType(newStreamConfig, vmt, profileData.VideoCodecData, PayloadType.dynamicVideo))
                {
                    return false;
                }

                //Set keyframe spacing
                IWMVideoMediaProps vmp = (IWMVideoMediaProps)newStreamConfig;
                vmp.SetMaxKeyFrameSpacing((long)8 * (long)Constants.TicksPerSec);  //PRI1: can the real data come from the header?

                //set bitrate, bufferwindow ...
                newStreamConfig.SetBitrate(vmt.VideoInfo.BitRate);
                newStreamConfig.SetBufferWindow(5000);

                //Add new streamconfig to profile
                try
                {
                    profile.AddStream(newStreamConfig);
                }
                catch { return false; }

                //Keep track of the stream number for writing stream samples
                videoStreamNum = nextStreamNum;
                nextStreamNum++;
            }
            else
            { 
                //Video MediaType is null, check the secondary profile info.  This supports slides in the video stream.
                if (profileData.VideoCodecGuid != Guid.Empty)
                {
                    //Create a stream config by enumerating installed codecs
                    IWMStreamConfig newStreamConfig = CreateStreamConfig(WMGuids.WMMEDIATYPE_Video, profileData.VideoCodecGuid, pm);
                    if (newStreamConfig == null)
                    {
                        return false;
                    }

                    //Get the media type from the stream config
                    byte[] codecData;
                    MediaTypeVideoInfo mt = (MediaTypeVideoInfo)GetStreamConfigMediaType(newStreamConfig, PayloadType.dynamicVideo, out codecData);

                    //Set the key properties from profileData
                    mt.VideoInfo.Source = new RECT();
                    mt.VideoInfo.Source.top = 0;
                    mt.VideoInfo.Source.left = 0;
                    mt.VideoInfo.Source.right = profileData.Width;
                    mt.VideoInfo.Source.bottom = profileData.Height;
                    mt.VideoInfo.Target = new RECT();
                    mt.VideoInfo.Target.top = 0;
                    mt.VideoInfo.Target.left = 0;
                    mt.VideoInfo.Target.right = profileData.Width;
                    mt.VideoInfo.Target.bottom = profileData.Height;
                    mt.VideoInfo.BitmapInfo.Height = profileData.Height;
                    mt.VideoInfo.BitmapInfo.Width = profileData.Width;
                    mt.VideoInfo.BitRate = profileData.BitRate;
                    mt.VideoInfo.AvgTimePerFrame = profileData.AvgTimePerFrame;

                    //Set the stream config media type
                    if (!SetStreamConfigMediaType(newStreamConfig, mt, codecData,PayloadType.dynamicVideo))
                        return false;

                    newStreamConfig.SetBitrate(profileData.BitRate);
                    newStreamConfig.SetBufferWindow(profileData.BufferWindow);
                    
                    //Here we also have to manually set the stream number
                    newStreamConfig.SetStreamNumber(nextStreamNum);

                    try
                    {
                        profile.AddStream(newStreamConfig);
                    }
                    catch
                    {
                        return false;
                    }

                    //Keep track of the stream number for writing stream samples
                    videoStreamNum = nextStreamNum;
                    nextStreamNum++;
                }
            }


            //Tell the writer to use this profile
            try
            {
                writer.SetProfile(profile);
                //this.DebugPrintProfile(profile);
            }
            catch { return false; }

            return true;

        }

        private bool SetStreamConfigMediaType(IWMStreamConfig sc, MediaType mt, byte[] codecData, PayloadType payload)
        {
            _WMMediaType wmmt;

            if (payload == PayloadType.dynamicVideo)
            {
                wmmt = ConvertVideoMediaType((MediaTypeVideoInfo)mt, codecData);
            }
            else if (payload == PayloadType.dynamicAudio)
            {
                wmmt = ConvertAudioMediaType((MediaTypeWaveFormatEx)mt, codecData);
            }
            else
                return false;

            IWMMediaProps mp = (IWMMediaProps)sc;
            bool ret = true;
            try
            {
                mp.SetMediaType(ref wmmt);
            }
            catch
            {
                ret = false;
            }
            finally
            {
                Marshal.FreeCoTaskMem(wmmt.pbFormat);
            }
            return ret;
        }

        /// <summary>
        /// Given a IWMStreamConfig for audio or video, return the MediaType from it.
        /// </summary>
        /// <param name="streamConfig"></param>
        /// <returns></returns>
        private MediaType GetStreamConfigMediaType(IWMStreamConfig streamConfig, PayloadType payload, out byte[] codecData)
        {
            //Get the MediaType
            IWMMediaProps mp = (IWMMediaProps)streamConfig;
            uint cMT = 0;
            mp.GetMediaType(IntPtr.Zero, ref cMT);
            IntPtr iMT = Marshal.AllocCoTaskMem((int)cMT);
            mp.GetMediaType(iMT, ref cMT);
            byte[] ba = new byte[cMT];
            Marshal.Copy(iMT, ba, 0, (int)cMT);
            BufferChunk bc = new BufferChunk(ba);

            MediaType mt = null;
            codecData = null;
            if (payload == PayloadType.dynamicVideo)
            {
                mt = new MediaTypeVideoInfo();
                ProfileUtility.ReconstituteBaseMediaType(mt, bc);
                ProfileUtility.ReconstituteVideoFormat((MediaTypeVideoInfo)mt, bc, out codecData);
            }
            else if (payload == PayloadType.dynamicAudio)
            {
                mt = new MediaTypeWaveFormatEx();
                ProfileUtility.ReconstituteBaseMediaType(mt, bc);
                ProfileUtility.ReconstituteAudioFormat((MediaTypeWaveFormatEx)mt, bc, out codecData);
            }
            return mt;
        }

        
        /// <summary>
        /// Enumerate the codecs of the specified majorType, and return a IWMStreamConfig from the first format
        /// of the first codec that supports the specified subType.  Return null if no match is found.
        /// </summary>
        /// <param name="guidMajorType"></param>
        /// <param name="guidSubType"></param>
        /// <param name="pm"></param>
        /// <returns></returns>
        private IWMStreamConfig CreateStreamConfig(Guid guidMajorType, Guid guidSubType, IWMProfileManager pm)
        {
            IWMCodecInfo3 ci3 = (IWMCodecInfo3)pm;
            uint cCodecs;
            GUID g = WMGuids.ToGUID(guidMajorType);
            ci3.GetCodecInfoCount(ref g, out cCodecs);

            for (uint i = 0; i < cCodecs; i++)
            {
                ////Retrieve the name for debugging purposes
                //uint cName = 0;
                //ci3.GetCodecName(ref g, i, IntPtr.Zero, ref cName);
                //IntPtr ipName = Marshal.AllocCoTaskMem((int)cName * 2);
                //ci3.GetCodecName(ref g, i, ipName, ref cName);
                //String s = Marshal.PtrToStringUni(ipName);
                //Debug.WriteLine("codec index=" + i.ToString() + ":" + s);
                //Marshal.FreeCoTaskMem(ipName);

                uint cFormat;
                ci3.GetCodecFormatCount(ref g, i, out cFormat);
                for (uint j = 0; j < cFormat; j++)
                {
                    IWMStreamConfig streamConfig;
                    ci3.GetCodecFormat(ref g, i, j, out streamConfig);

                    ////Show format descriptions for debugging purposes
                    //uint cDesc = 0; ;
                    //ci3.GetCodecFormatDesc(ref g, i, j, out streamConfig, IntPtr.Zero, ref cDesc);
                    //IntPtr ipDesc = Marshal.AllocCoTaskMem((int)cDesc * 2);
                    //ci3.GetCodecFormatDesc(ref g, i, j, out streamConfig, ipDesc, ref cDesc);
                    //String s2 = Marshal.PtrToStringUni(ipDesc);
                    //Debug.WriteLine("  format index=" + j.ToString() + ":" + s2);
                    //Marshal.FreeCoTaskMem(ipDesc);

                    //Get the IWMStreamConfig's MediaType
                    IWMMediaProps mp = (IWMMediaProps)streamConfig;
                    uint cMT = 0;
                    mp.GetMediaType(IntPtr.Zero, ref cMT);
                    IntPtr iMT = Marshal.AllocCoTaskMem((int)cMT);
                    mp.GetMediaType(iMT, ref cMT);
                    byte[] ba = new byte[cMT];
                    Marshal.Copy(iMT, ba, 0, (int)cMT);
                    BufferChunk bc = new BufferChunk(ba);
                    MediaType mt = new MediaType();
                    ProfileUtility.ReconstituteBaseMediaType(mt, bc);

                    if (mt.SubTypeAsGuid == guidSubType)
                    {
                        return streamConfig;
                    }
                }
            }
            return null;
        }


		/// <summary>
		/// The MediaType from the header for Screen codec streams seems to be missing a bunch of data.
		/// The critical one to fix was the Compression field in the BitmapInfo struct.
		/// Others don't seem to matter.
		/// </summary>
		/// <param name="vmt"></param>
		private void FixMTForScreenCodec(MediaTypeVideoInfo vmt)
		{
			if (vmt.SubType != SubType.MSS2)
				return;
			//Among the things that seem to be missing:
			//-VideoInfo.AverageTimePerFrame is zero
			//-Source/Target rects are all zeros
			//-bitcount is zero
			//-bitmapinfo.Compression is not set.

			//vmt.VideoInfo.Source.right = vmt.VideoInfo.BitmapInfo.Width;
			//vmt.VideoInfo.Source.bottom = vmt.VideoInfo.BitmapInfo.Height;
			//vmt.VideoInfo.Target.right = vmt.VideoInfo.BitmapInfo.Width;
			//vmt.VideoInfo.Target.bottom = vmt.VideoInfo.BitmapInfo.Height;
			//vmt.VideoInfo.BitmapInfo.BitCount = 24; //??
			//vmt.VideoInfo.AvgTimePerFrame = 666666;

			//This appears to be critical:
			vmt.VideoInfo.BitmapInfo.Compression = 0x3253534D; //hex MSS2

            //If this is zero the decompression won't work (reader returns no samples)
            if (vmt.VideoInfo.BitmapInfo.BitCount == 0)
            {
                //It seems a like a huge hack to assume a value, but not sure how else to proceed.
                vmt.VideoInfo.BitmapInfo.BitCount = 24;
            }
		}

		/// <summary>
		/// Iterate over writer inputs, holding on to the IWMInputMediaProps* and input numbers for each.
		/// Call this before ConfigAudio, and ConfigVideo methods.
		/// </summary>
		public void GetInputProps()
		{
			uint cInputs;
			writer.GetInputCount(out cInputs);
			GUID                guidInputType;
			IWMInputMediaProps  inputProps = null;
			for(uint i = 0; i < cInputs; i++ )
			{
				writer.GetInputProps( i, out inputProps );
				inputProps.GetType( out guidInputType );
				if( WMGuids.ToGuid(guidInputType) == WMGuids.WMMEDIATYPE_Audio )
				{
					audioProps = inputProps;
					audioInput = i;
				}
				else if( WMGuids.ToGuid(guidInputType) ==  WMGuids.WMMEDIATYPE_Video )
				{
					videoProps = inputProps;
					videoInput = i;
				}
				else if( WMGuids.ToGuid(guidInputType) == WMGuids.WMMEDIATYPE_Script )
				{
				}
				else
				{
					Debug.WriteLine( "Profile contains unrecognized media type." );
				}
			}

		}

		/// <summary>
		/// If samples are compressed (known as "Stream Samples") we want to SetInputProps to null for all inputs.
		/// Also find out the input numbers for audio and video for later.
		/// </summary>
		public void ConfigNullProps()
		{
			uint cInputs;
			writer.GetInputCount(out cInputs);
			GUID                guidInputType;
			IWMInputMediaProps  inputProps = null;

			for(uint i = 0; i < cInputs; i++ )
			{
				writer.GetInputProps( i, out inputProps );
				inputProps.GetType( out guidInputType );
				if( WMGuids.ToGuid(guidInputType) == WMGuids.WMMEDIATYPE_Audio )
				{
					audioInput = i;
				}
				else if( WMGuids.ToGuid(guidInputType) ==  WMGuids.WMMEDIATYPE_Video )
				{
					videoInput = i;
				}

				writer.SetInputProps(i,null);
			}
		}

		/// <summary>
		/// If samples are compressed, we need to put codec info into the header manually.
		/// </summary>
		public void SetCodecInfo(PayloadType payload)
		{
			// -IWMHeaderInfo3::AddCodecInfo for each stream
			// Video codec info:
			// type=WMT_CODECINFO_VIDEO;name=Windows Media Video V9;desc=;info=WMV3
			// Audio guessing..:
			// type=WMT_CODECINFO_AUDIO;name=Windows Media Audio V7;desc= 32 kbps, 32 kHz, stereo;info=a\001
			IWMHeaderInfo3 hi3 = (IWMHeaderInfo3)writer;	
			IntPtr name;
			IntPtr desc;
			IntPtr info;

			if (payload == PayloadType.dynamicAudio)
			{
				name = Marshal.StringToCoTaskMemUni("Windows Media Audio V7");
				desc = Marshal.StringToCoTaskMemUni(" 20 kbps, 22 kHz, stereo");
				byte[] binfo = {(byte)'a', 1};
				info = Marshal.AllocCoTaskMem(2);
				Marshal.Copy(binfo,0,info,2);
				hi3.AddCodecInfo(name,desc,WMT_CODEC_INFO_TYPE.WMT_CODECINFO_AUDIO,2,info);
				Marshal.FreeCoTaskMem(name);
				Marshal.FreeCoTaskMem(desc);
				Marshal.FreeCoTaskMem(info);
			}
			else if (payload == PayloadType.dynamicVideo)
			{
				name = Marshal.StringToCoTaskMemUni("Windows Media Video 9");
				//name = Marshal.StringToCoTaskMemUni("Windows Media Video 9 Screen");
				desc = Marshal.StringToCoTaskMemUni("");
				byte[] vbinfo = {(byte)'W', (byte)'M', (byte)'V', (byte)'3'};
				//byte[] vbinfo = {(byte)'M', (byte)'S', (byte)'S', (byte)'2'};
				info = Marshal.AllocCoTaskMem(4);
				Marshal.Copy(vbinfo,0,info,4);
				hi3.AddCodecInfo(name,desc,WMT_CODEC_INFO_TYPE.WMT_CODECINFO_VIDEO,4,info);
				Marshal.FreeCoTaskMem(name);
				Marshal.FreeCoTaskMem(desc);
				Marshal.FreeCoTaskMem(info);
			}
		}

		/// <summary>
		/// If samples are compressed, we need to put codec info into the header manually.
		/// This overload calls SetCodecInfo for both audio and video payloads.
		/// </summary>
		public void SetCodecInfo()
		{
			SetCodecInfo(PayloadType.dynamicVideo);
			SetCodecInfo(PayloadType.dynamicAudio);
		}

		/// <summary>
		/// Before writing uncompressed samples, call this method to configure the writer with the 
		/// uncompressed audio media type
		/// </summary>
		/// <param name="mt"></param>
		/// <returns></returns>
		public bool ConfigAudio(MediaTypeWaveFormatEx mt)
		{
			_WMMediaType _mt = ConvertAudioMediaType(mt);
			return ConfigAudio(_mt);
		}

		/// <summary>
		/// Before writing uncompressed samples, call this method to configure the writer with the 
		/// uncompressed audio media type
		/// </summary>
		public bool ConfigAudio(_WMMediaType mt)
		{
			if (audioProps == null)
			{
				Debug.WriteLine("Failed to configure audio: properties is null.");
				return false;
			}

			try
			{
				audioProps.SetMediaType(ref mt );
				writer.SetInputProps( audioInput, audioProps );
			}
			catch (Exception e)
			{
				Debug.WriteLine("Failed to set audio properties: " + e.ToString());
				return false;
			}
			return true;
		}



		/// <summary>
		/// Before writing uncompressed samples, call this method to configure the writer with the 
		/// uncompressed video media type.
		/// </summary>
		/// <param name="mt"></param>
		/// <returns></returns>
		public bool ConfigVideo(MediaTypeVideoInfo mt)
		{
            if (mt.SubType == SubType.RGB24) //this is a hack to repair a case that came up with Screen Streaming.
            {
                mt.VideoInfo.BitmapInfo.BitCount = 24;
                mt.VideoInfo.BitmapInfo.SizeImage = ((uint)mt.VideoInfo.BitmapInfo.Width * (uint)mt.VideoInfo.BitmapInfo.Height * 24) / 8;
            }
			_WMMediaType _mt = ConvertVideoMediaType(mt);
			return ConfigVideo(_mt);
		}

		/// <summary>
		/// Before writing uncompressed samples, call this method to configure the writer with the 
		/// uncompressed video media type.
		/// </summary>
		/// <param name="mt"></param>
		public bool ConfigVideo(_WMMediaType mt)
		{
			if (videoProps == null)
			{
				Debug.WriteLine("Failed to configure video: properties is null.");
				return false;
			}

			try
			{
				videoProps.SetMediaType( ref mt );
				writer.SetInputProps( videoInput, videoProps );
			}
			catch (Exception e)
			{
				Debug.WriteLine("Failed to set video properties: " + e.ToString());
				return false;
			}
			return true;
		}



		/// <summary>
		/// Signal that we are ready to start writing.  
		/// Call this method after configuring the profile and inputs, and before writing.
		/// </summary>
		/// <returns></returns>
		public bool Start()
		{
			try
			{
				writer.BeginWriting();
			}
			catch (Exception e)
			{
				Debug.WriteLine("Failed to start writing: " + e.ToString());
				return false;
			}
			return true;
		}


		/// <summary>
		/// Write an uncompressed audio sample.  Sample time is in ticks, relative to the start of the archive output.
		/// </summary>
		/// <param name="sampleSize"></param>
		/// <param name="inBuf"></param>
		/// <param name="sampleTime">in Ticks</param>
		/// <returns></returns>
		public bool WriteAudio(uint sampleSize, BufferChunk inBuf, ulong sampleTime)
		{
			INSSBuffer sample;
			IntPtr sampleBuf = IntPtr.Zero;

			//Debug.WriteLine("WMWriter.WriteAudio: time=" + sampleTime.ToString() + ";inputNum=" + audioInput);
			//return true;
			//	+ " size=" + sampleSize.ToString() +
			// " audio bytes " + inBuf[345].ToString() + " " +
			//	inBuf[346].ToString() + " " + inBuf[347].ToString() + " " + inBuf[348].ToString());

			try
			{
				lock(this)
				{
					writer.AllocateSample(sampleSize, out sample);
					sample.GetBuffer(out sampleBuf);
					Marshal.Copy(inBuf.Buffer,inBuf.Index,sampleBuf,(int)sampleSize);
					sample.SetLength(sampleSize);
					writer.WriteSample(audioInput,sampleTime,0,sample);
					//Debug.WriteLine("Wrote audio. time=" + sampleTime.ToString());
					Marshal.ReleaseComObject(sample);
				}
			}
			catch (Exception e)
			{
				//Debug.WriteLine("Exception while writing audio: " +
				//	"audioInput=" + videoInput.ToString() + 
				//	" sampleTime=" + sampleTime.ToString() + 
				//	" sampleSize=" + sampleSize.ToString());

				Debug.WriteLine("Exception while writing audio: " + e.ToString());
				//eventLog.WriteEntry("Exception while writing audio: " + e.ToString(), EventLogEntryType.Error, 1000);
				return false;
			}
			//Debug.WriteLine("Audio write succeeded " +
			//		"audioInput=" + videoInput.ToString() + 
			//		" sampleTime=" + sampleTime.ToString() + 
			//		" sampleSize=" + sampleSize.ToString());
			return true;
		}

		/// <summary>
		/// Write an uncompressed video sample.  Sample time is in ticks, relative to the start of the archive output.
		/// </summary>
		/// <param name="sampleSize"></param>
		/// <param name="inBuf"></param>
		/// <param name="sampleTime"></param>
		/// <returns></returns>
		public bool WriteVideo(uint sampleSize, BufferChunk inBuf, ulong sampleTime)
		{			
			INSSBuffer sample;
			IntPtr sampleBuf = IntPtr.Zero;

			//Debug.WriteLine("WMWriter.WriteVideo: time=" + sampleTime.ToString());
			//return true;
			//	+ " size=" + sampleSize.ToString() +
			//  " video bytes " + inBuf[345].ToString() + " " +
			//	inBuf[346].ToString() + " " + inBuf[347].ToString() + " " + inBuf[348].ToString());

			try
			{
				lock(this)
				{
					writer.AllocateSample(sampleSize, out sample);
					sample.GetBuffer(out sampleBuf);
					Marshal.Copy(inBuf.Buffer,inBuf.Index,sampleBuf,(int)sampleSize);
					sample.SetLength(sampleSize);
					writer.WriteSample(videoInput,sampleTime,0,(INSSBuffer)sample);
					//Debug.WriteLine("Wrote video. time=" + sampleTime.ToString());
					Marshal.ReleaseComObject(sample);
				}
			}
			catch (Exception)
			{
				//The most common cause of this seems to be failing to set or reset the media type correctly.
				Debug.WriteLine("Exception while writing video: " +
					"videoInput=" + videoInput.ToString() + 
					" sampleTime=" + sampleTime.ToString() + 
					" sampleSize=" + sampleSize.ToString());
				return false;
			}
			//Debug.WriteLine("Video write succeeded 	videoInput=" + videoInput.ToString() +
			//		" sampleTime=" + sampleTime.ToString() +
			//		" sampleSize=" + sampleSize.ToString());
			return true;
		}

		/// <summary>
		/// Write a compressed audio sample (aka: Stream Sample)
		/// </summary>
		/// <param name="sampleTime"></param>
		/// <param name="sampleSize"></param>
		/// <param name="buf"></param>
		/// <returns></returns>
		public bool WriteCompressedAudio(UInt64 sampleTime, uint sampleSize, byte[] buf)
		{
			//Note: WriteStreamSample params SampleSendTime and SampleDuration are unused.
			INSSBuffer sample;
			IntPtr sampleBuf = IntPtr.Zero;

			try
			{
				lock(this)
				{
					writer.AllocateSample(sampleSize, out sample);
					sample.GetBuffer(out sampleBuf);
					Marshal.Copy(buf,0,sampleBuf,(int)sampleSize);
					sample.SetLength(sampleSize);
					writerAdvanced.WriteStreamSample(audioStreamNum,sampleTime,0,0,0,(INSSBuffer)sample);
					Marshal.ReleaseComObject(sample);
					//Debug.WriteLine("Wrote audio: " +
					//	"Stream Number=" + audioStreamNum.ToString() + 
					//	" sampleTime=" + sampleTime.ToString() + 
					//	" sampleSize=" + sampleSize.ToString());
				}
			}
			catch (Exception e)
			{
				//The most common cause of this seems to be failing to set or reset the media type correctly.
				Debug.WriteLine("Exception while writing audio: " +
					"Stream Number=" + audioStreamNum.ToString() + 
					" sampleTime=" + sampleTime.ToString() + 
					" sampleSize=" + sampleSize.ToString() + e.ToString());
				return false;
			}

			return true;
		}


		/// <summary>
		/// Write a compressed video sample (aka: Stream Sample)
		/// </summary>
		/// <param name="sampleTime"></param>
		/// <param name="sampleSize"></param>
		/// <param name="buf"></param>
		/// <param name="keyframe"></param>
		/// <param name="discontinuity"></param>
		/// <returns></returns>
		public bool WriteCompressedVideo(UInt64 sampleTime, uint sampleSize, byte[] buf, bool keyframe, bool discontinuity)
		{
			//Note: WriteStreamSample params SampleSendTime and SampleDuration are unused.
			INSSBuffer sample;
			IntPtr sampleBuf = IntPtr.Zero;

			try
			{
				lock(this)
				{
					writer.AllocateSample(sampleSize, out sample);
					sample.GetBuffer(out sampleBuf);
					Marshal.Copy(buf,0,sampleBuf,(int)sampleSize);
					sample.SetLength(sampleSize);
					//Debug.WriteLine("length=" + sampleSize.ToString());
					uint flags  = 0;
					if (keyframe) //We should use the AM_SAMPLE_SPLICEPOINT value of the SampleFlags in AM_SAMPLE2_PROPERTIES to find keyframes..
					{
						flags = 1;
						//Debug.WriteLine("WriteCompressedVideo: keyframe.");
					}
					if (discontinuity)
					{
						flags += 2;
						//Debug.WriteLine("WriteCompressedVideo: discontinuity.");
					}
					writerAdvanced.WriteStreamSample(videoStreamNum,sampleTime,0,0,flags,(INSSBuffer)sample);
					Marshal.ReleaseComObject(sample);
                    //Debug.WriteLine("Wrote video: " +
                    //    "Stream number=" + videoStreamNum.ToString() +
                    //    " sampleTime=" + sampleTime.ToString() +
                    //    " sampleSize=" + sampleSize.ToString() +
                    //    " flags=" + flags.ToString());
				}
			}
			catch (Exception)
			{
				//The most common cause of this seems to be failing to set or reset the media type correctly.
				Debug.WriteLine("Exception while writing video: " +
					"Stream number=" + videoStreamNum.ToString() + 
					" sampleTime=" + sampleTime.ToString() + 
					" sampleSize=" + sampleSize.ToString());
				return false;
			}

			return true;
		}

		/// <summary>
		/// Signal that we are finished writing samples.
		/// </summary>
		public void Stop()
		{
			try
			{
				writer.EndWriting();
			}
			catch (Exception e)
			{
				Debug.WriteLine("Failed to stop: " + e.ToString());
			}
		}

        public bool SetFixedFrameRate(bool set)
        {
            if (writer != null)
            {
                if (videoInput == 0)
                {
                    GetInputProps();
                }

                if (videoInput != 0)
                {
                    IWMWriterAdvanced2 writerA2 = (IWMWriterAdvanced2)writer;

                    IntPtr ipName = Marshal.StringToCoTaskMemUni("FixedFrameRate");

                    //Watch out: WMT_TYPE_BOOL is 4 bytes.
                    WMT_ATTR_DATATYPE dtype = WMT_ATTR_DATATYPE.WMT_TYPE_BOOL;
                    ushort valLen = 4;
                    IntPtr ipValue = Marshal.AllocCoTaskMem(valLen);
                    Marshal.WriteInt32(ipValue, Convert.ToInt32(set));
                    //Marshal.WriteInt32(ipValue, 16843009);

                    //Hmm. There seems to be no difference no matter what I do here.
                    try
                    {
                        //writerA2.GetInputSetting(videoInput, ipName, out dtype, out ipValueOut, ref len);
                        writerA2.SetInputSetting(videoInput, ipName, dtype, ipValue, valLen);
                    }
                    catch (Exception e)
                    {
                        Debug.WriteLine(e.ToString());
                    }
                    finally
                    {
                        Marshal.FreeCoTaskMem(ipValue);
                        Marshal.FreeCoTaskMem(ipName);
                    }

                }
            }

            return true;
        }

		#endregion Public Methods

		#region Private Methods

		/// <summary>
		/// Return the name of a profile.
		/// </summary>
		/// <param name="profile"></param>
		/// <returns></returns>
		private static String GetProfileName(IWMProfile profile)
		{
			try
			{
				uint size = 0;
				profile.GetName(IntPtr.Zero,ref size);
				IntPtr buffer = Marshal.AllocCoTaskMem( (int)(2*(size+1)) );
				profile.GetName(buffer,ref size);
				String name = Marshal.PtrToStringAuto(buffer);
				Marshal.FreeCoTaskMem( buffer );
				return name;
			}
			catch (Exception e)
			{
				Debug.WriteLine("Failed to get profile name: " + e.ToString());
				return "";
			}
		}

		/// <summary>
		/// Convert profile indices and guids for some of the useful system profiles provided by the SDK
		/// </summary>
		/// <param name="index"></param>
		/// <returns></returns>
		private Guid ProfileIndexToGuid(uint index)
		{
			switch (index)
			{
				case 8:
					return WMGuids.WMProfile_V80_100Video;
				case 9:
					return WMGuids.WMProfile_V80_256Video;
				case 10:
					return WMGuids.WMProfile_V80_384Video;
				case 11:
					return WMGuids.WMProfile_V80_768Video;
				default: 
					return Guid.Empty;
			}
		}
	
		/// <summary>
		/// Load custom profile from a file or a string.  Ignore prxFile unless prxString is null or empty.
		/// </summary>
		/// <param name="prxFile"></param>
		/// <returns></returns>
		private IWMProfile LoadCustomProfile(String prxString, String prxFile)
		{
			IWMProfile pr;
			StreamReader stream;
			String s;
			if ((prxString==null) || (prxString==""))
			{
				try
				{
					stream = new StreamReader(prxFile);
				}
				catch (Exception e)
				{
					Debug.WriteLine("Failed to open profile file: " + e.ToString());
					return null;
				}
				s = stream.ReadToEnd();
			}
			else
			{
				s = prxString;
			}

			try
			{
				IntPtr profileStr = Marshal.StringToCoTaskMemUni(s);
				profileManager.LoadProfileByData(profileStr,out pr);
				Marshal.FreeCoTaskMem(profileStr);
				//Debug.WriteLine(s);
			}
			catch (Exception e)
			{
				Debug.WriteLine("Failed to load custom profile: " + s);
				Debug.WriteLine(e.ToString());
				return null;
			}
			return pr;
		}


		/// <summary>
		/// Copy the LST managed video MediaType to the Windows Media Interop type.  Use this overload when 
        /// compression data does not exist, or is not significant, such as when writing uncompressed samples.
		/// Notice that the IntPtr containing the format type should be freed by the caller when appropriate.
		/// </summary>
		/// <param name="mt"></param>
		/// <returns></returns>
		private _WMMediaType ConvertVideoMediaType(MediaTypeVideoInfo vi)
		{
			_WMMediaType wmt = new _WMMediaType();

			#region test settings
			// Basic video settings:
			//int w=320;
			//int h=240;
			//int fps=30;	

			// For RGB24:
			//ushort bpp=24;
			//uint comp=0; 
			//GUID stype = WMGuids.ToGUID(WMGuids.WMMEDIASUBTYPE_RGB24);

			// ..or for I420:
			//WORD bpp=12;
			//DWORD comp=0x30323449;
			//GUID stype= WMMEDIASUBTYPE_I420;

			// Settings for the video stream:
			// BITMAPINFOHEADER
			//  DWORD  biSize = size of the struct in bytes.. 40
			//	LONG   biWidth - Frame width
			//	LONG   biHeight	- height could be negative indicating top-down dib.
			//	WORD   biPlanes - must be 1.
			//	WORD   biBitCount 24 in our sample with RGB24
			//	DWORD  biCompression 0 for RGB
			//	DWORD  biSizeImage in bytes.. biWidth*biHeight*biBitCount/8
			//	LONG   biXPelsPerMeter 0
			//	LONG   biYPelsPerMeter 0; 
			//	DWORD  biClrUsed must be 0
			//	DWORD  biClrImportant 0
			//
			//	notes:
			//		biCompression may be a packed 'fourcc' code, for example I420 is 0x30323449, IYUV = 0x56555949...
			//		I420 and IYUV are identical formats.  They use 12 bits per pixel, and are planar,  comprised of
			//		nxm Y plane followed by n/2 x m/2 U and V planes.  Each plane is 8bits deep.

			//BitmapInfo bi = new BitmapInfo();
			//bi.Size=(uint)Marshal.SizeOf(bi);
			//bi.Width = w;
			//bi.Height = h;
			//bi.Planes = 1; //always 1.
			//bi.BitCount = bpp;
			//bi.Compression = comp; //RGB is zero.. uncompressed.
			//bi.SizeImage = (uint)(w * h * bpp / 8);
			//bi.XPelsPerMeter = 0;
			//bi.YPelsPerMeter = 0;
			//bi.ClrUsed = 0;
			//bi.ClrImportant = 0;

			// WMVIDEOINFOHEADER
			//  RECT  rcSource;
			//	RECT  rcTarget;
			//	DWORD  dwBitRate.. bps.. Width*Height*BitCount*Rate.. 320*240*24*29.93295=55172414
			//	DWORD  dwBitErrorRate zero in our sample.
			//	LONGLONG  AvgTimePerFrame in 100ns units.. 334080=10000*1000/29.93295
			//	BITMAPINFOHEADER  bmiHeader copy of the above struct.
			//VideoInfo vi = new VideoInfo();
			//vi.Source.left	= 0;
			//vi.Source.top	= 0;
			//vi.Source.bottom = bi.Height;
			//vi.Source.right	= bi.Width;
			//vi.Target		= vi.Source;
			//vi.BitRate		= (uint)(w * h * bpp * fps);
			//vi.BitErrorRate	= 0;
			//vi.AvgTimePerFrame = (UInt64) ((10000 * 1000) / fps);
			//vi.BitmapInfo = bi;


			// WM_MEDIA_TYPE
			//	GUID  majortype WMMEDIATYPE_Video
			//	GUID  subtype WMMEDIASUBTYPE_RGB24 in our sample
			//	BOOL  bFixedSizeSamples TRUE
			//	BOOL  bTemporalCompression FALSE
			//	ULONG  lSampleSize in bytes This was zero in our sample, but could be 320*240*24/8=230400
			//	GUID  formattype WMFORMAT_VideoInfo
			//	IUnknown*  pUnk NULL
			//	ULONG  cbFormat size of the WMVIDEOINFOHEADER 
			//	[size_is(cbFormat)] BYTE  *pbFormat pointer to the WMVIDEOINFOHEADER 
			#endregion
		
			//Put a format type in a IntPtr
			IntPtr viPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(vi.VideoInfo));
			Marshal.StructureToPtr(vi.VideoInfo,viPtr,true);

			//Fill in the base type.
			wmt.majortype = WMGuids.ToGUID(vi.MajorTypeAsGuid);
			wmt.subtype = WMGuids.ToGUID(vi.SubTypeAsGuid);
			wmt.bFixedSizeSamples = vi.FixedSizeSamples?1:0;
			wmt.bTemporalCompression = vi.TemporalCompression?1:0;
			wmt.lSampleSize = 0; 
			wmt.formattype = WMGuids.ToGUID(vi.FormatTypeAsGuid);
			wmt.pUnk = null;
			wmt.cbFormat = (uint)Marshal.SizeOf(vi.VideoInfo);
			wmt.pbFormat = viPtr; //IntPtr to format type

			return wmt;
		}

		/// <summary>
		/// Copy the LST managed video MediaType to the Windows Media Interop type
		/// </summary>
		/// This version also appends the codecdata to the BitmapInfo header.  This is required when
		/// configuring profiles prior to writing stream samples.
		/// Notice that the IntPtr containing the format type should be freed by the caller when appropriate.
		/// <param name="mt"></param>
		/// <returns></returns>
		private _WMMediaType ConvertVideoMediaType(MediaTypeVideoInfo vi, byte[] compressionData)
		{
			//Fix the size field to account for the compression data.  Unlike audio, it should be 
            //the sum of the size of the BitmapInfo and the compression data, and it is most likely not correct.
			vi.VideoInfo.BitmapInfo.Size = (uint)Marshal.SizeOf(vi.VideoInfo.BitmapInfo) + (uint)compressionData.Length;

            //Watch out for Marshal.SizeOf.  It returns the next multiple of 4 bytes (on 32-bit).
            //The true size of BitmapInfo is 40 bytes, and the true size of VideoInfo is 88, so we're ok here.

			//Put the format type and compressionData into a IntPtr.
			IntPtr viPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(vi.VideoInfo) + compressionData.Length);
			Marshal.StructureToPtr(vi.VideoInfo,viPtr,false);
			int ptr = Marshal.SizeOf(vi.VideoInfo) + viPtr.ToInt32();
			Marshal.Copy(compressionData,0,(IntPtr)ptr,compressionData.Length);

			//Create and fill in the interop type
			_WMMediaType wmt = new _WMMediaType();
			wmt.majortype = WMGuids.ToGUID(vi.MajorTypeAsGuid);
			wmt.subtype = WMGuids.ToGUID(vi.SubTypeAsGuid);
			wmt.bFixedSizeSamples = vi.FixedSizeSamples?1:0;
			wmt.bTemporalCompression = vi.TemporalCompression?1:0;
			wmt.lSampleSize = 0; 
			wmt.formattype = WMGuids.ToGUID(vi.FormatTypeAsGuid);
			wmt.pUnk = null;
			wmt.cbFormat = (uint)Marshal.SizeOf(vi.VideoInfo) + (uint)compressionData.Length;
			wmt.pbFormat = viPtr; //pointer to the format type

			return wmt;
		}


        /// <summary>
        /// Copy the LST managed audio MediaType to the Windows Media Interop type
        /// </summary>
        /// <param name="mt"></param>
        /// <returns></returns>
        /// Note that the IntPtr to the format type should be freed by the caller when appropriate.
        private _WMMediaType ConvertAudioMediaType(MediaTypeWaveFormatEx wfex, byte[] compressionData)
        {
            //Note that Marshal.SizeOf rounds up to the next 4-byte word.  In our case that screws up
            //the placement of the compression data.  We should be calculating the true size of WaveFormatEx, 
            //but for now, just hard-code it.
            int wfsz_marshal = Marshal.SizeOf(wfex.WaveFormatEx);
            int wfsz_true = 18;

            //Note that unlike video, the wfex.WaveFormatEx.Size value is exactly compressionData.Length.

            //Put the format type and compressionData into a IntPtr.
            IntPtr wfPtr = Marshal.AllocCoTaskMem(wfsz_marshal + compressionData.Length);
            Marshal.StructureToPtr(wfex.WaveFormatEx, wfPtr, false);
            int ptr = wfsz_true + wfPtr.ToInt32();
            Marshal.Copy(compressionData, 0, (IntPtr)ptr, compressionData.Length);

            _WMMediaType wmt = new _WMMediaType();
            wmt.majortype = WMGuids.ToGUID(wfex.MajorTypeAsGuid); 
            wmt.subtype = WMGuids.ToGUID(wfex.SubTypeAsGuid); 
            wmt.bFixedSizeSamples = wfex.FixedSizeSamples ? 1 : 0;
            wmt.bTemporalCompression = wfex.TemporalCompression ? 1 : 0;
            wmt.lSampleSize = (uint)wfex.SampleSize; 
            wmt.formattype = WMGuids.ToGUID(wfex.FormatTypeAsGuid);
            wmt.pUnk = null;
            wmt.cbFormat = (uint)wfsz_true + (uint)compressionData.Length;
            wmt.pbFormat = wfPtr;

            return wmt;
        }


		/// <summary>
		/// Copy the LST managed audio MediaType to the Windows Media Interop type in the case
        /// where there is no compression data (when writing uncompressed samples)
		/// Note that the IntPtr to the format type should be freed by the caller when appropriate.
		/// </summary>
		private _WMMediaType ConvertAudioMediaType(MediaTypeWaveFormatEx wfex)
		{
			_WMMediaType wmt = new _WMMediaType();

			IntPtr wfexPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(wfex.WaveFormatEx));
			Marshal.StructureToPtr(wfex.WaveFormatEx,wfexPtr,true);

			wmt.majortype = WMGuids.ToGUID(wfex.MajorTypeAsGuid); 
			wmt.subtype = WMGuids.ToGUID(wfex.SubTypeAsGuid); 
			wmt.bFixedSizeSamples = wfex.FixedSizeSamples?1:0;
			wmt.bTemporalCompression = wfex.TemporalCompression?1:0;
			wmt.lSampleSize	= (uint)wfex.SampleSize; 
			wmt.formattype = WMGuids.ToGUID(wfex.FormatTypeAsGuid);
			wmt.pUnk = null;
			wmt.cbFormat = (uint)Marshal.SizeOf( wfex.WaveFormatEx ) + wfex.WaveFormatEx.Size;
			wmt.pbFormat = wfexPtr;

			#region test code
			//				WaveFormatEx wfex = new WaveFormatEx();
			//
			//				wfex.FormatTag = 1; //1==WAVE_FORMAT_PCM
			//				wfex.Channels = 1;
			//				wfex.SamplesPerSec = 16000;
			//				wfex.AvgBytesPerSec =  32000;
			//				wfex.BlockAlign = 2;
			//				wfex.BitsPerSample = 16;
			//				wfex.Size = 0;

			//try
			//{
			//  Used GetMediaType to sanity check the managed structs:
			//uint size = 0;
			//audioProps.GetMediaType(IntPtr.Zero,ref size);
			//IntPtr mtPtr = Marshal.AllocCoTaskMem((int)size);
			//audioProps.GetMediaType(mtPtr,ref size);
			//_WMMediaType mt2 = (_WMMediaType)Marshal.PtrToStructure(mtPtr,typeof(_WMMediaType));
			//WMMediaType.WaveFormatEx wfex2 = (WMMediaType.WaveFormatEx)Marshal.PtrToStructure(mt2.pbFormat,typeof(WMMediaType.WaveFormatEx));
			//  Examine here.
			//Marshal.StructureToPtr(mt,mtPtr,true);
			//audioProps.SetMediaType( mtPtr );
			//}
			//catch (Exception e)
			//{
			//	Debug.WriteLine("Failed to set audio properties: " + e.ToString());
			//	return wmt;
			//}	
			#endregion

			return wmt;
		}

		/// <summary>
		/// Get the string (XML) representation from a configured IWMProfile and print to the debug console.
		/// </summary>
		/// <param name="profile"></param>
		internal void DebugPrintProfile(IWMProfile profile)
		{
			//Create a IWMProfileManager
			IWMProfileManager pm;
			WMFSDKFunctions.WMCreateProfileManager(out pm);

			uint pLen = 0;
			pm.SaveProfile(profile, IntPtr.Zero, ref pLen);
			IntPtr ipps = Marshal.AllocCoTaskMem((int)pLen * 2);
			pm.SaveProfile(profile, ipps, ref pLen);
			String sps = Marshal.PtrToStringUni(ipps);
			Debug.WriteLine(sps);
			Marshal.FreeCoTaskMem(ipps);
		}

        internal void DebugPrintVideoProps()
        {
            if (this.videoProps == null)
            {
                Debug.WriteLine("Video Props is null.");
                return;
            }
            uint cmt = 0;
            _WMMediaType wmMediaType;

            videoProps.GetMediaType(IntPtr.Zero, ref cmt);
            IntPtr imt = Marshal.AllocCoTaskMem((int)cmt);
            videoProps.GetMediaType(imt, ref cmt);

            wmMediaType = (_WMMediaType)Marshal.PtrToStructure(imt, typeof(_WMMediaType));
            byte[] ba = new byte[cmt];
            Marshal.Copy(imt, ba, 0, (int)cmt);
            BufferChunk bc = new BufferChunk(ba);
            byte[] codecData;

            MediaTypeVideoInfo mt = new MediaTypeVideoInfo();
            ProfileUtility.ReconstituteBaseMediaType(mt, bc);
            ProfileUtility.ReconstituteVideoFormat(mt, bc, out codecData);

            ProfileUtility.DebugPrintBaseMediaType(mt);
            ProfileUtility.DebugPrintVideoFormat(mt);

            Marshal.FreeCoTaskMem(imt);
        }		

		#endregion Private Methods
	
		#region Unused and Test Code

        ////For testing: unused
        //public bool ConfigSlideStreamVideo(MediaTypeVideoInfo new_mt)
        //{
        //    if (this.videoProps == null)
        //    {
        //        throw (new Exception("videoprops is null."));
        //    }

        //    uint cmt = 0;
        //    this.videoProps.GetMediaType(IntPtr.Zero, ref cmt);
        //    IntPtr imt = Marshal.AllocCoTaskMem((int)cmt);
        //    videoProps.GetMediaType(imt, ref cmt);
        //    byte[] bmt = new byte[cmt];
        //    Marshal.Copy(imt, bmt, 0, (int)cmt);
        //    BufferChunk bc = new BufferChunk(bmt);

        //    MediaTypeVideoInfo mt = new MediaTypeVideoInfo();
        //    ProfileUtility.ReconstituteBaseMediaType(mt, bc);
        //    byte[] codecData;
        //    ProfileUtility.ReconstituteVideoFormat(mt, bc, out codecData);

        //    //set rcSource, rcTarget
        //    mt.VideoInfo.Source = new RECT();
        //    mt.VideoInfo.Source.top = 0;
        //    mt.VideoInfo.Source.left = 0;
        //    mt.VideoInfo.Source.right = 320;
        //    mt.VideoInfo.Source.bottom = 240;
        //    mt.VideoInfo.Target = new RECT();
        //    mt.VideoInfo.Target.top = 0;
        //    mt.VideoInfo.Target.left = 0;
        //    mt.VideoInfo.Target.right = 320;
        //    mt.VideoInfo.Target.bottom = 240;

        //    //set width, height
        //    mt.VideoInfo.BitmapInfo.Height = 240;
        //    mt.VideoInfo.BitmapInfo.Width = 320;

        //    ProfileUtility.DebugPrintBaseMediaType(mt);
        //    ProfileUtility.DebugPrintVideoFormat(mt);
        //    ProfileUtility.DebugPrintBaseMediaType(new_mt);
        //    ProfileUtility.DebugPrintVideoFormat(new_mt);

        //    //note there could be other things we need to set in the input props such as
        //    // maximum image size.
        //    //mt.VideoInfo.BitmapInfo.SizeImage = 230400;
        //    //mt.VideoInfo.AvgTimePerFrame = 0;

        //    //_WMMediaType wmmt = ConvertVideoMediaType(mt, codecData);
        //    _WMMediaType wmmt = ConvertVideoMediaType(new_mt, codecData);

        //    try
        //    {
        //        videoProps.SetMediaType(ref wmmt);
        //        writer.SetInputProps(this.videoInput, videoProps);
        //    }
        //    catch (Exception e)
        //    {
        //        Debug.WriteLine("Failed to reset media type: " + e.ToString());
        //        return false;
        //    }
        //    return true;
        //}
        ///// <summary>
        ///// Create a profile completely from scratch and configure it with the data in the byte[].
        ///// Used this method to write raw video for streams using the Screen codec.  It also worked for normal CXP video.
        ///// </summary>
        ///// <param name="bmt"></param>
        //public bool ConfigNewVideoProfile(byte[] bmt)
        //{

        //    //Create a IWMProfileManager
        //    IWMProfileManager pm;
        //    uint hr = WMFSDKFunctions.WMCreateProfileManager(out pm);

        //    //create a empty profile
        //    IWMProfile profile;
        //    pm.CreateEmptyProfile(WMT_VERSION.WMT_VER_9_0, out profile);

        //    //create a empty stream
        //    IWMStreamConfig newStreamConfig;
        //    GUID g = WMGuids.ToGUID(WMGuids.WMMEDIATYPE_Video);
        //    profile.CreateNewStream(ref g, out newStreamConfig);

        //    //Convert the byte[] to a MediaTypeVideoInfo
        //    MediaTypeVideoInfo vmt = new MediaTypeVideoInfo();
        //    BufferChunk bc = new BufferChunk(bmt);
        //    ProfileUtility.ReconstituteBaseMediaType((MediaType)vmt, bc);
        //    byte[] compressionData = null;
        //    ProfileUtility.ReconstituteVideoFormat(vmt, bc, out compressionData);

        //    //"fix" vmt so that it will work when the screen codec is used..
        //    FixMTForScreenCodec(vmt);

        //    //Build the profile string from this media type for debugging purposes
        //    //String prxString = ProfileUtility.VideoMediaTypeToProfile(vmt,compressionData);
        //    //Debug.WriteLine(prxString);

        //    //Configure media type of the new StreamConfig
        //    _WMMediaType wmmt = this.ConvertVideoMediaType(vmt,compressionData);
        //    IWMMediaProps mp = (IWMMediaProps)newStreamConfig;
        //    mp.SetMediaType(ref wmmt);
        //    Marshal.FreeCoTaskMem(wmmt.pbFormat);

        //    //Set keyframe spacing
        //    IWMVideoMediaProps vmp = (IWMVideoMediaProps)newStreamConfig;
        //    vmp.SetMaxKeyFrameSpacing(8*Constants.TicksPerSec);  //PRI1: can the real data come from the header?

        //    //set bitrate, bufferwindow ...
        //    newStreamConfig.SetBitrate(vmt.VideoInfo.BitRate);  
        //    newStreamConfig.SetBufferWindow(5000);

        //    //Things that could be wrong here still:
        //    //may not have text strings such as name/description set
        //    //may not have decoder complexity set
        //    //maxkeyframespacing is made up (may not be correct)
        //    //quality is not set

        //    //Add new streamconfig to profile
        //    profile.AddStream(newStreamConfig);

        //    //Print the text representation for debugging purposes
        //    DebugPrintProfile(profile);

        //    //Tell the writer to use this profile
        //    writer.SetProfile(profile);

        //    //Initialize input and stream numbers.
        //    audioInput = videoInput = 0;
        //    audioProps = videoProps = null;
        //    videoStreamNum = 1;

        //    return true;
        //}


		//		private void EnumerateCodecs()
		//		{
		//			IWMCodecInfo3 ci3 = (IWMCodecInfo3)profileManager;
		//			uint cCodecs;
		//
		//			//Video codecs
		//			GUID g =  WMGuids.ToGUID(WMGuids.WMMEDIATYPE_Video);
		//			ci3.GetCodecInfoCount(ref g,out cCodecs);
		//			uint cchar = 0;
		//			IntPtr name;
		//			String sname;
		//			uint cformat = 0;
		//			IWMStreamConfig streamConfig;
		//			IWMMediaProps mediaProps;
		//			for (uint i=0;i<cCodecs;i++)
		//			{
		//				ci3.GetCodecName(ref g,i,IntPtr.Zero,ref cchar);
		//				name = Marshal.AllocCoTaskMem((int)((cchar*2) + 1));
		//				ci3.GetCodecName(ref g, i, name, ref cchar);
		//				sname = Marshal.PtrToStringUni(name);
		//				//Console.WriteLine("VideoCodec:" + sname);
		//				ci3.GetCodecFormatCount(ref g,i,out cformat);
		//				for (uint j=0;j<cformat;j++)
		//				{
		//					ci3.GetCodecFormat(ref g, i, j, out streamConfig);
		//					mediaProps = (IWMMediaProps)streamConfig;
		//					uint cmt = 0;
		//					MediaTypeVideoInfo mt = new MediaTypeVideoInfo();
		//					IntPtr imt;
		//					mediaProps.GetMediaType(IntPtr.Zero, ref cmt);
		//					imt = Marshal.AllocCoTaskMem((int)cmt);
		//					mediaProps.GetMediaType(imt, ref cmt);
		//					byte[] bmt = new byte[cmt];
		//					Marshal.Copy(imt,bmt,0,(int)cmt);
		//					BufferChunk bc = new BufferChunk(bmt);
		//					ReconstituteBaseMediaType((MediaType)mt,bc);
		//					ReconstituteVideoFormat(mt,bc);
		//				}
		//			}
		//		}

		/// <summary>
		/// Prepare the network writer.
		/// </summary>
		/// <param name="port"></param>
		/// <param name="maxClients"></param>
		/// <returns></returns>
		//		public bool ConfigNet(uint port, uint maxClients)
		//		{
		//			if ((writerAdvanced == null) || (netSink == null))
		//			{
		//				Debug.WriteLine("WriterAdvanced and NetSink must exist before calling ConfigNet");
		//				return false;
		//			}
		//
		//			try
		//			{
		//				netSink.SetNetworkProtocol(WMT_NET_PROTOCOL.WMT_PROTOCOL_HTTP);
		//
		//				netSink.Open(ref port);
		//			
		//				uint size = 0;
		//				netSink.GetHostURL(IntPtr.Zero,ref size);
		//				IntPtr buf = Marshal.AllocCoTaskMem( (int)(2*(size+1)) );
		//				netSink.GetHostURL(buf,ref size);
		//				String url = Marshal.PtrToStringAuto(buf);
		//				Marshal.FreeCoTaskMem( buf );
		//				Debug.WriteLine("Connect to:" + url);
		//
		//				netSink.SetMaximumClients(maxClients);
		//				writerAdvanced.AddSink(netSink);
		//			}
		//			catch (Exception e)
		//			{
		//				eventLog.WriteEntry("Failed to configure network: " + e.ToString(), EventLogEntryType.Error, 1000);
		//				Debug.WriteLine("Failed to configure network: " + e.ToString());
		//				return false;
		//			}
		//			return true;
		//		}

		//		/// <summary>
		//		/// Unused: Create a profile for the native CXP data
		//		/// </summary>
		//		/// <param name="vbr"></param>
		//		/// <param name="abr"></param>
		//		private void ConfigCXPProfile(uint vbr, uint abr)
		//		{
		//		
		//			IWMProfileManager pm;
		//			IWMProfile profile;
		//			IWMStreamConfig streamConfig;
		//			uint hr = WMFSDKFunctions.WMCreateProfileManager(out pm);
		//			pm.CreateEmptyProfile(WMT_VERSION.WMT_VER_9_0, out profile);
		//
		//			//Video:
		//			GUID videoGUID = WMGuids.ToGUID(WMGuids.WMMEDIATYPE_Video);
		//			profile.CreateNewStream(ref videoGUID,out streamConfig);
		//			uint ui;
		//			streamConfig.GetBitrate(out ui);
		//			//AddStream fails if SetBitrate is changed??  Does it only accept certain values?
		//			streamConfig.SetBitrate(ui);
		//			streamConfig.SetBufferWindow(2000); //2 seconds
		//			//ushort snum;
		//			streamConfig.GetStreamNumber(out videoStreamNum);
		//			profile.AddStream(streamConfig);
		//
		//			//Audio:
		//			GUID audioGUID = WMGuids.ToGUID(WMGuids.WMMEDIATYPE_Audio);
		//			profile.CreateNewStream(ref audioGUID,out streamConfig);
		//			streamConfig.GetBitrate(out ui);
		//			streamConfig.SetBitrate(ui);
		//			streamConfig.SetBufferWindow(2000); //2 seconds
		//			streamConfig.GetStreamNumber(out audioStreamNum);
		//			profile.AddStream(streamConfig);
		//
		//			writer.SetProfile(profile);
		//
		//		}

		//		/// <summary>
		//		/// Return a list of all the system profiles.
		//		/// </summary>
		//		/// <returns></returns>
		//		public static String QuerySystemProfiles()
		//		{
		//			IWMProfileManager pm;
		//			IWMProfileManager2 pm2;
		//
		//			uint hr = WMFSDKFunctions.WMCreateProfileManager(out pm);
		//			pm2 = (IWMProfileManager2)pm;
		//			pm2.SetSystemProfileVersion(WMT_VERSION.WMT_VER_7_0);
		//			pm2 = null;
		//
		//			uint pCount;
		//			pm.GetSystemProfileCount(out pCount);
		//
		//			IWMProfile profile;
		//
		//			StringBuilder sb = new StringBuilder(500);
		//			sb.Append("System Profile count: " + pCount.ToString() + "\n\r");
		//			String name;
		//			for (uint i=0;i<pCount;i++)
		//			{
		//				pm.LoadSystemProfile(i,out profile);
		//				name = GetProfileName(profile);
		//				sb.Append((i+1).ToString() + "  " + name + "\n\r");
		//			}
		//			
		//			return(sb.ToString());
		//
		//		}

		//		public void SetAudioCodecInfo()
		//		{
		//			IWMHeaderInfo3 hi3 = (IWMHeaderInfo3)writer;
		//			IntPtr name = Marshal.StringToCoTaskMemUni("Windows Media Audio V7");
		//			IntPtr desc = Marshal.StringToCoTaskMemUni(" 20 kbps, 22 kHz, stereo");
		//			byte[] binfo = {(byte)'a', 1};
		//			IntPtr info = Marshal.AllocCoTaskMem(2);
		//			Marshal.Copy(binfo,0,info,2);
		//			hi3.AddCodecInfo(name,desc,WMT_CODEC_INFO_TYPE.WMT_CODECINFO_AUDIO,2,info);
		//			Marshal.FreeCoTaskMem(name);
		//			Marshal.FreeCoTaskMem(desc);
		//			Marshal.FreeCoTaskMem(info);
		//		}

		//		/// <summary>
		//		/// Hardcode audio config for testing.
		//		/// </summary>
		//		/// <returns></returns>
		//		public bool ConfigAudio()
		//		{
		//			//make up some media types for testing
		//			
		//			WAVEFORMATEX wfex = new WAVEFORMATEX();
		//			
		//			wfex.FormatTag = 1; //1==WAVE_FORMAT_PCM
		//			wfex.Channels = 1;
		//			wfex.SamplesPerSec = 16000;
		//			wfex.AvgBytesPerSec =  32000;
		//			wfex.BlockAlign = 2;
		//			wfex.BitsPerSample = 16;
		//			wfex.Size = 0;
		//
		//			IntPtr wfexPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(wfex));
		//			Marshal.StructureToPtr(wfex,wfexPtr,true);
		//
		//			_WMMediaType mt = new _WMMediaType();
		//			mt.majortype			= WMGuids.ToGUID(WMGuids.WMMEDIATYPE_Audio);
		//			mt.subtype				= WMGuids.ToGUID(WMGuids.WMMEDIASUBTYPE_PCM);
		//			mt.bFixedSizeSamples	= 1; //true
		//			mt.bTemporalCompression = 0; //false
		//			mt.lSampleSize			= 2;
		//			mt.formattype			= WMGuids.ToGUID(WMGuids.WMFORMAT_WaveFormatEx);  //This is the only value permitted.
		//			mt.pUnk					= null;
		//			mt.cbFormat				= (uint)Marshal.SizeOf( wfex ) + wfex.Size;
		//			mt.pbFormat				= wfexPtr;
		//
		//			try
		//			{
		//				//  Used GetMediaType to sanity check the managed structs:
		//				//uint size = 0;
		//				//audioProps.GetMediaType(IntPtr.Zero,ref size);
		//				//IntPtr mtPtr = Marshal.AllocCoTaskMem((int)size);
		//				//audioProps.GetMediaType(mtPtr,ref size);
		//				//_WMMediaType mt2 = (_WMMediaType)Marshal.PtrToStructure(mtPtr,typeof(_WMMediaType));
		//				//WMMediaType.WaveFormatEx wfex2 = (WMMediaType.WaveFormatEx)Marshal.PtrToStructure(mt2.pbFormat,typeof(WMMediaType.WaveFormatEx));
		//				//  Examine here.
		//				//Marshal.StructureToPtr(mt,mtPtr,true);
		//				//audioProps.SetMediaType( mtPtr );
		//			}
		//			catch (Exception e)
		//			{
		//				//eventLog.WriteEntry("Failed to set audio properties: " + e.ToString(), EventLogEntryType.Error, 1000);
		//				Debug.WriteLine("Failed to set audio properties: " + e.ToString());
		//				return false;
		//			}
		//
		//			bool ret = ConfigAudio(mt);
		//			
		//			Marshal.FreeCoTaskMem(wfexPtr);
		//			return ret;
		//
		//		}

		//		/// <summary>
		//		/// Hardcode video config for testing.
		//		/// </summary>
		//		/// <returns></returns>
		//		public bool ConfigVideo()
		//		{
		//			// Basic video settings:
		//			int w=320;
		//			int h=240;
		//			int fps=30;	
		//	
		//			// For RGB24:
		//			ushort bpp=24;
		//			uint comp=0; 
		//			GUID stype = WMGuids.ToGUID(WMGuids.WMMEDIASUBTYPE_RGB24);
		//
		//			// ..or for I420:
		//			//WORD bpp=12;
		//			//DWORD comp=0x30323449;
		//			//GUID stype= WMMEDIASUBTYPE_I420;
		//
		//			// Settings for the video stream:
		//			// BITMAPINFOHEADER
		//			//  DWORD  biSize = size of the struct in bytes.. 40
		//			//	LONG   biWidth - Frame width
		//			//	LONG   biHeight	- height could be negative indicating top-down dib.
		//			//	WORD   biPlanes - must be 1.
		//			//	WORD   biBitCount 24 in our sample with RGB24
		//			//	DWORD  biCompression 0 for RGB
		//			//	DWORD  biSizeImage in bytes.. biWidth*biHeight*biBitCount/8
		//			//	LONG   biXPelsPerMeter 0
		//			//	LONG   biYPelsPerMeter 0; 
		//			//	DWORD  biClrUsed must be 0
		//			//	DWORD  biClrImportant 0
		//			//
		//			//	notes:
		//			//		biCompression may be a packed 'fourcc' code, for example I420 is 0x30323449, IYUV = 0x56555949...
		//			//		I420 and IYUV are identical formats.  They use 12 bits per pixel, and are planar,  comprised of
		//			//		nxm Y plane followed by n/2 x m/2 U and V planes.  Each plane is 8bits deep.
		//			
		//			BITMAPINFOHEADER bi = new BITMAPINFOHEADER();
		//			bi.Size=(uint)Marshal.SizeOf(bi);
		//			bi.Width = w;
		//			bi.Height = h;
		//			bi.Planes = 1; //always 1.
		//			bi.BitCount = bpp;
		//			bi.Compression = comp; //RGB is zero.. uncompressed.
		//			bi.SizeImage = (uint)(w * h * bpp / 8);
		//			bi.XPelsPerMeter = 0;
		//			bi.YPelsPerMeter = 0;
		//			bi.ClrUsed = 0;
		//			bi.ClrImportant = 0;
		//
		//			// WMVIDEOINFOHEADER
		//			//  RECT  rcSource;
		//			//	RECT  rcTarget;
		//			//	DWORD  dwBitRate.. bps.. Width*Height*BitCount*Rate.. 320*240*24*29.93295=55172414
		//			//	DWORD  dwBitErrorRate zero in our sample.
		//			//	LONGLONG  AvgTimePerFrame in 100ns units.. 334080=10000*1000/29.93295
		//			//	BITMAPINFOHEADER  bmiHeader copy of the above struct.
		//			VIDEOINFOHEADER vi = new VIDEOINFOHEADER();
		//			vi.Source.left	= 0;
		//			vi.Source.top	= 0;
		//			vi.Source.bottom = bi.Height;
		//			vi.Source.right	= bi.Width;
		//			vi.Target		= vi.Source;
		//			vi.BitRate		= (uint)(w * h * bpp * fps);
		//			vi.BitErrorRate	= 0;
		//			vi.AvgTimePerFrame = (UInt64) ((10000 * 1000) / fps);
		//			vi.BitmapInfo = bi;
		//			
		//			IntPtr viPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(vi));
		//			Marshal.StructureToPtr(vi,viPtr,true);
		//
		//			// WM_MEDIA_TYPE
		//			//	GUID  majortype WMMEDIATYPE_Video
		//			//	GUID  subtype WMMEDIASUBTYPE_RGB24 in our sample
		//			//	BOOL  bFixedSizeSamples TRUE
		//			//	BOOL  bTemporalCompression FALSE
		//			//	ULONG  lSampleSize in bytes This was zero in our sample, but could be 320*240*24/8=230400
		//			//	GUID  formattype WMFORMAT_VideoInfo
		//			//	IUnknown*  pUnk NULL
		//			//	ULONG  cbFormat size of the WMVIDEOINFOHEADER 
		//			//	[size_is(cbFormat)] BYTE  *pbFormat pointer to the WMVIDEOINFOHEADER 
		//
		//			//Note WM_MEDIA_TYPE is the same as Directshow's AM_MEDIA_TYPE.
		//			//WM_MEDIA_TYPE   mt;
		//			_WMMediaType mt = new _WMMediaType();
		//			mt.majortype = WMGuids.ToGUID(WMGuids.WMMEDIATYPE_Video);
		//			mt.subtype = stype;
		//			mt.bFixedSizeSamples = 1;
		//			mt.bTemporalCompression = 0;
		//			//mt.lSampleSize = w * h * bpp / 8;  // this was zero in avinetwrite!
		//			mt.lSampleSize = 0; //hmm.  Don't think it matters??
		//			mt.formattype = WMGuids.ToGUID(WMGuids.WMFORMAT_VideoInfo);
		//			mt.pUnk = null;
		//			mt.cbFormat = (uint)Marshal.SizeOf(vi);
		//			mt.pbFormat = viPtr;
		//
		//			bool ret = ConfigVideo(mt);
		//			
		//			Marshal.FreeCoTaskMem(viPtr);
		//			return ret;
		//		}

		//		/// <summary>
		//		/// Write a script command into the stream
		//		/// </summary>
		//		/// <param name="type"></param>
		//		/// <param name="script"></param>
		//		/// <param name="packScript">Use both bytes of the unicode script string.  This is 
		//		/// used to more efficiently transmit Base64 encoded data in a script.</param>
		//		/// <returns></returns>
		//		/// The writer expects two null terminated WCHARs (two bytes per character).
		//		public bool SendScript(String type, String script, bool packScript)
		//		{
		//			IntPtr typePtr, scriptPtr, bufPtr;
		//			byte[] sampleBuf;
		//			INSSBuffer sample;
		//			ulong curTime;
		//			uint typesz, scriptsz, nulls;
		//
		//			//ScriptStreamNumber == 0 means there is no script stream in the profile.
		//			if (scriptStreamNumber == 0) 
		//			{
		//				return false;
		//			}
		//
		//			if (packScript)
		//			{
		//				typesz = (uint)(2 * (type.Length + 1));
		//				//To make the script looks like unicode, need to terminate with two nulls, and 
		//				//the total length needs to be an even number.
		//				scriptsz = (uint)script.Length;
		//				if (scriptsz%2 == 0)
		//				{
		//					nulls=2;
		//				} 
		//				else 
		//				{
		//					nulls=3;
		//				}
		//				sampleBuf = new byte[typesz+scriptsz+nulls];
		//				typePtr = Marshal.StringToCoTaskMemUni(type);
		//				scriptPtr = Marshal.StringToCoTaskMemAnsi(script);
		//				Marshal.Copy(typePtr,sampleBuf,0,(int)typesz);
		//				Marshal.Copy(scriptPtr,sampleBuf,(int)typesz,(int)scriptsz);
		//				for(uint i=typesz+scriptsz+nulls-1;i>=typesz+scriptsz;i--)
		//				{
		//					sampleBuf[i] = 0;
		//				}
		//				scriptsz += nulls;
		//				Marshal.FreeCoTaskMem(typePtr);
		//				Marshal.FreeCoTaskMem(scriptPtr);
		//			}
		//			else
		//			{
		//				//Marshal both strings as unicode.
		//				typesz = (uint)(2 * (type.Length + 1));
		//				scriptsz = (uint)(2 * (script.Length + 1));
		//				sampleBuf = new byte[typesz+scriptsz];
		//				typePtr = Marshal.StringToCoTaskMemUni(type);
		//				scriptPtr = Marshal.StringToCoTaskMemUni(script);
		//				Marshal.Copy(typePtr,sampleBuf,0,(int)typesz);
		//				Marshal.Copy(scriptPtr,sampleBuf,(int)typesz,(int)scriptsz);
		//				Marshal.FreeCoTaskMem(typePtr);
		//				Marshal.FreeCoTaskMem(scriptPtr);
		//			}
		//
		//			try
		//			{
		//				lock(this)
		//				{
		//					writer.AllocateSample((typesz+scriptsz),out sample);
		//					sample.GetBuffer(out bufPtr);
		//					Marshal.Copy(sampleBuf,0,bufPtr,(int)(typesz+scriptsz));
		//					// Let the writer tell us what time it wants to use to avoid
		//					// rebuffering and other nastiness.  Can this cause a loss of sync issue?
		//					writerAdvanced.GetWriterTime(out curTime);
		//					writerAdvanced.WriteStreamSample(scriptStreamNumber,curTime,0,0,0,sample);
		//					Marshal.ReleaseComObject(sample);
		//				}
		//			}
		//			catch (Exception e)
		//			{
		//				//eventLog.WriteEntry("Failed to write script: " + e.ToString(), EventLogEntryType.Error, 1000);
		//				Debug.WriteLine("Failed to write script: " + e.ToString());
		//				return false;
		//			}
		//			return true;
		//		}
		//

		//		/// <summary>
		//		/// Event handler for MediaBuffer OnSampleReady event
		//		/// </summary>
		//		/// <param name="sea"></param>
		//		//private void ReceiveSample(SampleEventArgs sea)
		//		private void ReceiveSample()
		//		{
		//			if (writeFailed)
		//			{
		//				return;
		//			}
		//
		////			TimeSpan ts = new TimeSpan((long)sea.Time);
		////			if (sea.Type == PayloadType.dynamicVideo)
		////			{
		////
		////				if (!WriteVideo((uint)sea.Buffer.Length,sea.Buffer,sea.Time))
		////				{
		////					writeFailed = true;
		////				}
		////				//Debug.WriteLine("WMWriter.ReceiveSample len=" + sea.Buffer.Length.ToString() +
		////				//	" time=" + ts.TotalSeconds + " type=video");
		////			}
		////			else if (sea.Type == PayloadType.dynamicAudio)
		////			{
		////				if (!WriteAudio((uint)sea.Buffer.Length,sea.Buffer,sea.Time))
		////				{
		////					writeFailed = true;
		////				}
		////				//Debug.WriteLine("WMWriter.ReceiveSample len=" + sea.Buffer.Length.ToString() +
		////				//	" time=" + ts.TotalSeconds + " type=audio");
		////			}
		//		}

 
#endregion Unused and Test Code

    }
}
