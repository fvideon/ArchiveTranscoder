using System;
using System.Runtime.InteropServices;
using MSR.LST.MDShow;


// Note: 
// WMVideo Encoder DMO (a.k.a. Windows Media 7) uses the WMV1 fourCC
// WMVideo9 Encoder DMO (a.k.a. Windows Media 9) uses the WMV3 fourCC
// WMVideo Advanced Encoder DMO (a.k.a. Windows Media 10 / 9.5 SDK) uses the WMVA fourCC

namespace MSR.LST.MDShow
{

    public class MediaType
    {
        #region Constructors
        public MediaType(){}
        public MediaType(_AMMediaType mt)
        {
            this.mt = mt;

            FixedSizeSamples = mt.bFixedSizeSamples;
            TemporalCompression = mt.bTemporalCompression;
            SampleSize = (int)mt.lSampleSize;
            MajorType = MajorTypeGuidToEnum(mt.majortype);
            SubType = SubTypeGuidToEnum(mt.subtype);
            FormatType = FormatTypeGuidToEnum(mt.formattype);
        }

        public MediaType(IntPtr pMT)
        {
            _AMMediaType am = new _AMMediaType();
            _AMMediaType mt = (_AMMediaType)System.Runtime.InteropServices.Marshal.PtrToStructure(pMT, am.GetType());
            this.mt = mt;

            FixedSizeSamples = mt.bFixedSizeSamples;
            TemporalCompression = mt.bTemporalCompression;
            SampleSize = (int)mt.lSampleSize;
            MajorType = MajorTypeGuidToEnum(mt.majortype);
            SubType = SubTypeGuidToEnum(mt.subtype);
            FormatType = FormatTypeGuidToEnum(mt.formattype);
        }

        #endregion

        #region Private Properties
        public _AMMediaType mt;
        #endregion

        #region Public Properties
        public MajorType MajorType;
        public SubType SubType;
        public FormatType FormatType;
        public bool FixedSizeSamples;
        public bool TemporalCompression;
        public int SampleSize;
        public Guid MajorTypeAsGuid
        {
            get
            {
                return mt.majortype;
            }
        }
        public Guid SubTypeAsGuid
        {
            get
            {
                return mt.subtype;
            }
        }
        public Guid FormatTypeAsGuid
        {
            get
            {
                return mt.formattype;
            }
        }
        #endregion

        #region Private Methods
        public static MajorType MajorTypeGuidToEnum(Guid guid)
        {
            if(guid == MEDIATYPE_Video )
                return MajorType.Video;
            if(guid == MEDIATYPE_Audio )
                return MajorType.Audio;
            if(guid == MEDIATYPE_Text )
                return MajorType.Text;
            if(guid == MEDIATYPE_Midi )
                return MajorType.Midi;
            if(guid == MEDIATYPE_Stream )
                return MajorType.Stream;
            if(guid == MEDIATYPE_Interleaved )
                return MajorType.Interleaved;
            if(guid == MEDIATYPE_File )
                return MajorType.File;
            if(guid == MEDIATYPE_ScriptCommand )
                return MajorType.ScriptCommand;
            if(guid == MEDIATYPE_AUXLine21Data )
                return MajorType.AUXLine21Data;
            if(guid == MEDIATYPE_VBI )
                return MajorType.VBI;
            if(guid == MEDIATYPE_Timecode )
                return MajorType.Timecode;
            if(guid == MEDIATYPE_LMRT )
                return MajorType.LMRT;
            if(guid == MEDIATYPE_URL_STREAM )
                return MajorType.UrlStream;

            throw new ArgumentException("Unrecognized Guid for MajorType");
        }
        public static Guid MajorTypeEnumToGuid(MajorType mt)
        {
            if (mt == MajorType.Video)
                return MEDIATYPE_Video;
            if (mt == MajorType.Audio )
                return MEDIATYPE_Audio;
            if (mt == MajorType.Text )
                return MEDIATYPE_Text;
            if (mt == MajorType.Midi )
                return MEDIATYPE_Midi;
            if (mt == MajorType.Stream )
                return MEDIATYPE_Stream;
            if (mt == MajorType.Interleaved )
                return MEDIATYPE_Interleaved;
            if (mt == MajorType.File )
                return MEDIATYPE_File;
            if (mt == MajorType.ScriptCommand )
                return MEDIATYPE_ScriptCommand;
            if (mt == MajorType.AUXLine21Data )
                return MEDIATYPE_AUXLine21Data;
            if (mt == MajorType.VBI )
                return MEDIATYPE_VBI;
            if (mt == MajorType.Timecode )
                return MEDIATYPE_Timecode;
            if (mt == MajorType.LMRT )
                return MEDIATYPE_LMRT;
            if (mt == MajorType.UrlStream )
                return MEDIATYPE_URL_STREAM;

            throw new ArgumentException("Unrecognized Guid for MajorType");
        }

        // SubTypeEnumToGuid
        public static SubType SubTypeGuidToEnum(Guid guid)
        {
            if(guid == WMMEDIASUBTYPE_WMV1 )return SubType.WMV1;
            if(guid == WMMEDIASUBTYPE_WMV2 )return SubType.WMV2;
            if(guid == WMMEDIASUBTYPE_WMV3 )return SubType.WMV3;
            if(guid == WMMEDIASUBTYPE_WMVP )return SubType.WMVP;
            if(guid == WMMEDIASUBTYPE_WMVA )return SubType.WMVA;
            if(guid == WMMEDIASUBTYPE_WVP2 )return SubType.WVP2;
            if(guid == MEDIASUBTYPE_I420) return SubType.I420;
            if(guid == MEDIASUBTYPE_CLPL ) return SubType.CLPL;
            if(guid == MEDIASUBTYPE_YUYV ) return SubType.YUYV;
            if(guid == MEDIASUBTYPE_IYUV ) return SubType.IYUV;
            if(guid == MEDIASUBTYPE_YVU9 ) return SubType.YVU9;
            if(guid == MEDIASUBTYPE_Y411 ) return SubType.Y411 ;
            if(guid == MEDIASUBTYPE_Y41P ) return SubType.Y41P;
            if(guid == MEDIASUBTYPE_YUY2 ) return SubType.YUY2;
            if(guid == MEDIASUBTYPE_YVYU ) return SubType.YVYU;
            if(guid == MEDIASUBTYPE_UYVY ) return SubType.UYVY;
            if(guid == MEDIASUBTYPE_Y211 ) return SubType.Y211;
            if(guid == MEDIASUBTYPE_YV12 ) return SubType.YV12;
            if(guid == MEDIASUBTYPE_CLJR ) return SubType.CLJR;
            if(guid == MEDIASUBTYPE_IF09 ) return SubType.IF09;
            if(guid == MEDIASUBTYPE_CPLA ) return SubType.CPLA;
            if(guid == MEDIASUBTYPE_MJPG0) return SubType.MJPG0;
            if(guid == MEDIASUBTYPE_TVMJ ) return SubType.TVMJ;
            if(guid == MEDIASUBTYPE_WAKE ) return SubType.WAKE;
            if(guid == MEDIASUBTYPE_CFCC ) return SubType.CFCC;
            if(guid == MEDIASUBTYPE_IJPG ) return SubType.IJPG;
            if(guid == MEDIASUBTYPE_Plum ) return SubType.Plum;
            if(guid == MEDIASUBTYPE_DVCS ) return SubType.DVCS;
            if(guid == MEDIASUBTYPE_DVSD ) return SubType.DVSD;
            if(guid == MEDIASUBTYPE_MDVF ) return SubType.MDVF;
            if(guid == MEDIASUBTYPE_RGB1 ) return SubType.RGB1;
            if(guid == MEDIASUBTYPE_RGB4 ) return SubType.RGB4;
            if(guid == MEDIASUBTYPE_RGB8 ) return SubType.RGB8;
            if(guid == MEDIASUBTYPE_RGB565 ) return SubType.RGB565;
            if(guid == MEDIASUBTYPE_RGB555 ) return SubType.RGB555;
            if(guid == MEDIASUBTYPE_RGB24 )   return SubType.RGB24;
            if(guid == MEDIASUBTYPE_RGB32 )   return SubType.RGB32;
            if(guid == MEDIASUBTYPE_ARGB1555 )    return SubType.ARGB1555;
            if(guid == MEDIASUBTYPE_ARGB4444 ) return SubType.ARGB4444;
            if(guid == MEDIASUBTYPE_ARGB32 ) return SubType.ARGB32;
            if(guid == MEDIASUBTYPE_AYUV ) return SubType.AYUV;
            if(guid == MEDIASUBTYPE_AI44 ) return SubType.AI44;
            if(guid == MEDIASUBTYPE_IA44 ) return SubType.IA44;
            if(guid == MEDIASUBTYPE_Overlay ) return SubType.Overlay;
            if(guid == MEDIASUBTYPE_MPEGPacket ) return SubType.MPEGPacket;
            if(guid == MEDIASUBTYPE_MPEG1Payload ) return SubType.MPEG1Payload;
            if(guid == MEDIASUBTYPE_MPEG1AudioPayload ) return SubType.MPEG1AudioPayload;
            if(guid == MEDIASUBTYPE_MPEG1SystemStream ) return SubType.MPEG1SystemStream;
            if(guid == MEDIASUBTYPE_MPEG1System ) return SubType.MPEG1System;
            if(guid == MEDIASUBTYPE_MPEG1VideoCD ) return SubType.MPEG1VideoCD;
            if(guid == MEDIASUBTYPE_MPEG1Video ) return SubType.MPEG1Video;
            if(guid == MEDIASUBTYPE_MPEG1Audio ) return SubType.MPEG1Audio;
            if(guid == MEDIASUBTYPE_Avi )
                return SubType.Avi;
            if(guid == MEDIASUBTYPE_Asf )
                return SubType.Asf;
            if(guid == MEDIASUBTYPE_QTMovie )
                return SubType.QTMovie;
            if(guid == MEDIASUBTYPE_QTRpza )
                return SubType.QTRpza;
            if(guid == MEDIASUBTYPE_QTSmc )
                return SubType.QTSmc;
            if(guid == MEDIASUBTYPE_QTRle )
                return SubType.QTRle;
            if(guid == MEDIASUBTYPE_QTJpeg )
                return SubType.QTJpeg;
            if(guid == MEDIASUBTYPE_PCMAudio_Obsolete )
                return SubType.PCMAudio_Obsolete;
            if(guid == MEDIASUBTYPE_PCM )
                return SubType.PCM;
            if(guid == MEDIASUBTYPE_WAVE )
                return SubType.WAVE;
            if(guid == MEDIASUBTYPE_AU )
                return SubType.AU;
            if(guid == MEDIASUBTYPE_AIFF )
                return SubType.AIFF;
            if(guid == MEDIASUBTYPE_dvsd )
                return SubType.dvsd;
            if(guid == MEDIASUBTYPE_dvhd )
                return SubType.dvhd;
            if(guid == MEDIASUBTYPE_dvsl )
                return SubType.dvsl;
            if(guid == MEDIASUBTYPE_Line21_BytePair )
                return SubType.Line21_BytePair;
            if(guid == MEDIASUBTYPE_Line21_GOPPacket )
                return SubType.Line21_GOPPacket;
            if(guid == MEDIASUBTYPE_Line21_VBIRawData )
                return SubType.Line21_VBIRawData;
            if(guid == MEDIASUBTYPE_TELETEXT )
                return SubType.TELETEXT;
            if(guid == MEDIASUBTYPE_DRM_Audio )
                return SubType.DRM_Audio;
            if(guid == MEDIASUBTYPE_IEEE_FLOAT )
                return SubType.IEEE_FLOAT;
            if(guid == MEDIASUBTYPE_DOLBY_AC3_SPDIF )
                return SubType.DOLBY_AC3_SPDIF;
            if(guid == MEDIASUBTYPE_RAW_SPORT )
                return SubType.RAW_SPORT;
            if(guid == MEDIASUBTYPE_SPDIF_TAG_241h )
                return SubType.SPDIF_TAG_241h;
            if(guid == MEDIASUBTYPE_DssVideo )
                return SubType.DssVideo;
            if(guid == MEDIASUBTYPE_DssAudio )
                return SubType.DssAudio;
            if(guid == MEDIASUBTYPE_VPVideo )
                return SubType.VPVideo;
            if(guid == MEDIASUBTYPE_VPVBI )
                return SubType.VPVBI;
            if(guid == WMMEDIASUBTYPE_MP43) return SubType.MP43;
            if(guid == WMMEDIASUBTYPE_mp43) return SubType.mp43;
            if(guid == WMMEDIASUBTYPE_MP4S ) return SubType.MP4S;
            if(guid == WMMEDIASUBTYPE_mp4s ) return SubType.mp4s;
			if(guid == WMMEDIASUBTYPE_MSS1 )
				return SubType.MSS1;
			if(guid == WMMEDIASUBTYPE_MSS2 )
				return SubType.MSS2;
			if(guid == WMMEDIASUBTYPE_PCM )
                return SubType.WMPCM;
            if(guid == WMMEDIASUBTYPE_DRM )
                return SubType.DRM;
            if(guid == WMMEDIASUBTYPE_WMAudioV7 )
                return SubType.WMAudioV7;
            if(guid == WMMEDIASUBTYPE_WMAudioV2 )
                return SubType.WMAudioV2;
            if(guid == WMMEDIASUBTYPE_ACELPnet )
                return SubType.ACELPnet;
            if (guid == MEDIASUBTYPE_UNDOC_Y422 )
                return SubType.UNDOC_Y422;

            return SubType.Unknown;
        }
        
        public static Guid SubTypeEnumToGuid(SubType st)
        {
            if(st == SubType.I420) return MEDIASUBTYPE_I420;
            if(st == SubType.WMV1) return WMMEDIASUBTYPE_WMV1;
            if(st == SubType.WMV2) return WMMEDIASUBTYPE_WMV2;
            if(st == SubType.WMV3) return WMMEDIASUBTYPE_WMV3;
            if(st == SubType.WMVP) return WMMEDIASUBTYPE_WMVP;
            if(st == SubType.WMVA) return WMMEDIASUBTYPE_WMVA;
            if(st == SubType.WVP2) return WMMEDIASUBTYPE_WVP2;
            if(st == SubType.UNDOC_Y422 )
                return MEDIASUBTYPE_UNDOC_Y422;
            if(st == SubType.CLPL)
                return MEDIASUBTYPE_CLPL;
            if(st == SubType.YUYV)
                return MEDIASUBTYPE_YUYV;
            if(st == SubType.IYUV)
                return MEDIASUBTYPE_IYUV;
            if(st == SubType.YVU9 )
                return MEDIASUBTYPE_YVU9;
            if(st == SubType.Y411 )
                return MEDIASUBTYPE_Y411;
            if(st == SubType.Y41P )
                return MEDIASUBTYPE_Y41P;
            if(st == SubType.YUY2 )
                return MEDIASUBTYPE_YUY2;
            if(st == SubType.YVYU )
                return MEDIASUBTYPE_YVYU;
            if  ( st == SubType.UYVY ) 
                return MEDIASUBTYPE_UYVY;
            if  ( st == SubType.Y211 ) 
                return MEDIASUBTYPE_Y211;
            if  ( st == SubType.YV12 ) 
                return MEDIASUBTYPE_YV12;
            if  ( st == SubType.CLJR ) 
                return MEDIASUBTYPE_CLJR;
            if  ( st == SubType.IF09 ) 
                return MEDIASUBTYPE_IF09;
            if  ( st == SubType.CPLA ) 
                return MEDIASUBTYPE_CPLA;
            if  ( st == SubType.MJPG0 ) 
                return MEDIASUBTYPE_MJPG0;
            if  ( st == SubType.TVMJ ) 
                return MEDIASUBTYPE_TVMJ;
            if  ( st == SubType.WAKE ) 
                return MEDIASUBTYPE_WAKE;
            if  ( st == SubType.CFCC ) 
                return MEDIASUBTYPE_CFCC;
            if  ( st == SubType.IJPG ) 
                return MEDIASUBTYPE_IJPG;
            if  ( st == SubType.Plum ) 
                return MEDIASUBTYPE_Plum;
            if  ( st == SubType.DVCS ) 
                return MEDIASUBTYPE_DVCS;
            if  ( st == SubType.DVSD ) 
                return MEDIASUBTYPE_DVSD;
            if  ( st == SubType.MDVF ) 
                return MEDIASUBTYPE_MDVF;
            if  ( st == SubType.RGB1 ) 
                return MEDIASUBTYPE_RGB1;
            if  ( st == SubType.RGB4 ) 
                return MEDIASUBTYPE_RGB4;
            if  ( st == SubType.RGB8 ) 
                return MEDIASUBTYPE_RGB8;
            if  ( st == SubType.RGB565 ) 
                return MEDIASUBTYPE_RGB565;
            if  ( st == SubType.RGB555 ) 
                return MEDIASUBTYPE_RGB555;
            if  ( st == SubType.RGB24 ) 
                return MEDIASUBTYPE_RGB24;
            if  ( st == SubType.RGB32 ) 
                return MEDIASUBTYPE_RGB32;
            if  ( st == SubType.ARGB1555 ) 
                return MEDIASUBTYPE_ARGB1555;
            if  ( st == SubType.ARGB4444 ) 
                return MEDIASUBTYPE_ARGB4444;
            if  ( st == SubType.ARGB32 ) 
                return MEDIASUBTYPE_ARGB32;
            if  ( st == SubType.AYUV ) 
                return MEDIASUBTYPE_AYUV;
            if  ( st == SubType.AI44 ) 
                return MEDIASUBTYPE_AI44;
            if  ( st == SubType.IA44 ) 
                return MEDIASUBTYPE_IA44;
            if  ( st == SubType.Overlay ) 
                return MEDIASUBTYPE_Overlay;
            if  ( st == SubType.MPEGPacket ) 
                return MEDIASUBTYPE_MPEGPacket;
            if  ( st == SubType.MPEG1Payload ) 
                return MEDIASUBTYPE_MPEG1Payload;
            if  ( st == SubType.MPEG1AudioPayload ) 
                return MEDIASUBTYPE_MPEG1AudioPayload;
            if  ( st == SubType.MPEG1SystemStream ) 
                return MEDIASUBTYPE_MPEG1SystemStream;
            if  ( st == SubType.MPEG1System ) 
                return MEDIASUBTYPE_MPEG1System;
            if  ( st == SubType.MPEG1VideoCD ) 
                return MEDIASUBTYPE_MPEG1VideoCD;
            if  ( st == SubType.MPEG1Video ) 
                return MEDIASUBTYPE_MPEG1Video;
            if  ( st == SubType.MPEG1Audio ) 
                return MEDIASUBTYPE_MPEG1Audio;
            if  ( st == SubType.Avi ) 
                return MEDIASUBTYPE_Avi;
            if  ( st == SubType.Asf ) 
                return MEDIASUBTYPE_Asf;
            if  ( st == SubType.QTMovie ) 
                return MEDIASUBTYPE_QTMovie;
            if  ( st == SubType.QTRpza ) 
                return MEDIASUBTYPE_QTRpza;
            if  ( st == SubType.QTSmc ) 
                return MEDIASUBTYPE_QTSmc;
            if  ( st == SubType.QTRle ) 
                return MEDIASUBTYPE_QTRle;
            if  ( st == SubType.QTJpeg ) 
                return MEDIASUBTYPE_QTJpeg;
            if  ( st == SubType.PCMAudio_Obsolete ) 
                return MEDIASUBTYPE_PCMAudio_Obsolete;
            if  ( st == SubType.PCM ) 
                return MEDIASUBTYPE_PCM;
            if  ( st == SubType.WAVE ) 
                return MEDIASUBTYPE_WAVE;
            if  ( st == SubType.AU ) 
                return MEDIASUBTYPE_AU;
            if  ( st == SubType.AIFF ) 
                return MEDIASUBTYPE_AIFF;
            if  ( st == SubType.dvsd ) 
                return MEDIASUBTYPE_dvsd;
            if  ( st == SubType.dvhd ) 
                return MEDIASUBTYPE_dvhd;
            if  ( st == SubType.dvsl ) 
                return MEDIASUBTYPE_dvsl;
            if  ( st == SubType.Line21_BytePair ) 
                return MEDIASUBTYPE_Line21_BytePair;
            if  ( st == SubType.Line21_GOPPacket ) 
                return MEDIASUBTYPE_Line21_GOPPacket;
            if  ( st == SubType.Line21_VBIRawData ) 
                return MEDIASUBTYPE_Line21_VBIRawData;
            if  ( st == SubType.TELETEXT ) 
                return MEDIASUBTYPE_TELETEXT;
            if  ( st == SubType.DRM_Audio ) 
                return MEDIASUBTYPE_DRM_Audio;
            if  ( st == SubType.IEEE_FLOAT ) 
                return MEDIASUBTYPE_IEEE_FLOAT;
            if  ( st == SubType.DOLBY_AC3_SPDIF ) 
                return MEDIASUBTYPE_DOLBY_AC3_SPDIF;
            if  ( st == SubType.RAW_SPORT ) 
                return MEDIASUBTYPE_RAW_SPORT;
            if  ( st == SubType.SPDIF_TAG_241h ) 
                return MEDIASUBTYPE_SPDIF_TAG_241h;
            if  ( st == SubType.DssVideo ) 
                return MEDIASUBTYPE_DssVideo;
            if  ( st == SubType.DssAudio ) 
                return MEDIASUBTYPE_DssAudio;
            if  ( st == SubType.VPVideo ) 
                return MEDIASUBTYPE_VPVideo;
            if  ( st == SubType.VPVBI ) 
                return MEDIASUBTYPE_VPVBI;
            if  ( st == SubType.MP43 ) return WMMEDIASUBTYPE_MP43;
            if  ( st == SubType.mp43 ) return WMMEDIASUBTYPE_mp43;
            if  ( st == SubType.MP4S ) return WMMEDIASUBTYPE_MP4S;
            if  ( st == SubType.mp4s ) return WMMEDIASUBTYPE_mp4s;
            if  ( st == SubType.MSS1 ) 
                return WMMEDIASUBTYPE_MSS1;
            if  ( st == SubType.WMPCM ) 
                return WMMEDIASUBTYPE_PCM;
            if  ( st == SubType.DRM ) 
                return WMMEDIASUBTYPE_DRM;
            if  ( st == SubType.WMAudioV7 ) 
                return WMMEDIASUBTYPE_WMAudioV7;
            if  ( st == SubType.WMAudioV2 ) 
                return WMMEDIASUBTYPE_WMAudioV2;
            if  ( st == SubType.ACELPnet)
                return WMMEDIASUBTYPE_ACELPnet;

            return Guid.Empty;
        }


        // FormatTypeEnumToGuid
        public static FormatType FormatTypeGuidToEnum(Guid guid)
        {

            if(guid == FORMAT_None )
                return FormatType.None;
            if(guid == FORMAT_VideoInfo )
                return FormatType.VideoInfo;
            if(guid == FORMAT_VideoInfo2 )
                return FormatType.VideoInfo2;
            if(guid == FORMAT_WaveFormatEx )
                return FormatType.WaveFormatEx;
            if(guid == FORMAT_MPEGVideo )
                return FormatType.MPEGVideo;
            if(guid == FORMAT_MPEGStreams )
                return FormatType.MPEGStreams;
            if(guid == FORMAT_DvInfo )
                return FormatType.DvInfo;

            throw new ArgumentException("Unknown Guid for FormatType");

        }

        public static Guid FormatTypeEnumToGuid(FormatType ft)
        {

            if(ft == FormatType.VideoInfo )
                return FORMAT_VideoInfo;
            if(ft ==  FormatType.None )
                return FORMAT_None;
            if(ft ==  FormatType.VideoInfo )
                return FORMAT_VideoInfo;
            if(ft ==  FormatType.VideoInfo2 )
                return FORMAT_VideoInfo2;
            if(ft ==  FormatType.WaveFormatEx )
                return FORMAT_WaveFormatEx;
            if(ft ==  FormatType.MPEGVideo )
                return FORMAT_MPEGVideo;
            if(ft ==  FormatType.MPEGStreams )
                return FORMAT_MPEGStreams;
            if(ft ==  FormatType.DvInfo )
                return FORMAT_DvInfo;

            throw new ArgumentException("Unknown Guid for FormatType");

        }

        internal VIDEOINFOHEADER GetVideoInfo (IntPtr pbFormat, uint cbFormat)
        {
            return (VIDEOINFOHEADER)Marshal.PtrToStructure(pbFormat, typeof(VIDEOINFOHEADER));
        }
        internal WAVEFORMATEX GetWaveFormatEx(IntPtr pbFormat, uint cbFormat)
        {
            return (WAVEFORMATEX)Marshal.PtrToStructure(pbFormat, typeof(WAVEFORMATEX));
        }
        #endregion

        #region Public Methods
        public MediaTypeVideoInfo ToMediaTypeVideoInfo()
        {
            if (FormatType != FormatType.VideoInfo)
                throw new ArgumentException("Invalid FormatType: must by VideoInfo");

            return new MediaTypeVideoInfo(mt);
        }
        public MediaTypeWaveFormatEx ToMediaTypeWaveFormatEx()
        {
            if (FormatType != FormatType.WaveFormatEx)
                throw new ArgumentException("Invalid FormatType: must by WaveFormatEx");

            return new MediaTypeWaveFormatEx(mt);
        }
        public virtual _AMMediaType ToAMMediaType()
        {
            return mt;
        }

        #endregion
        
        #region Major Types

        public static Guid MEDIATYPE_Video = new Guid(0x73646976, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIATYPE_Audio = new Guid(0x73647561, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIATYPE_Text = new Guid(0x73747874, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIATYPE_Midi = new Guid(0x7364696D, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIATYPE_Stream = new Guid(0xe436eb83, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIATYPE_Interleaved = new Guid(0x73766169, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIATYPE_File = new Guid(0x656c6966, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIATYPE_ScriptCommand = new Guid(0x73636d64, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIATYPE_AUXLine21Data = new Guid(0x670aea80, 0x3a82, 0x11d0, 0xb7, 0x9b, 0x0, 0xaa, 0x0, 0x37, 0x67, 0xa7);
        public static Guid MEDIATYPE_VBI = new Guid(0xf72a76e1, 0xeb0a, 0x11d0, 0xac, 0xe4, 0x00, 0x00, 0xc0, 0xcc, 0x16, 0xba);
        public static Guid MEDIATYPE_Timecode = new Guid(0x482dee3, 0x7817, 0x11cf, 0x8a, 0x3, 0x0, 0xaa, 0x0, 0x6e, 0xcb, 0x65);
        public static Guid MEDIATYPE_LMRT = new Guid(0x74726c6d, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIATYPE_URL_STREAM = new Guid(0x736c7275, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        
        
        #endregion
        
        #region Sub Types

        public static Guid MEDIASUBTYPE_RGB1                = new Guid(0xe436eb78, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_RGB4                = new Guid(0xe436eb79, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_RGB8                = new Guid(0xe436eb7a, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_RGB565              = new Guid(0xe436eb7b, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_RGB555              = new Guid(0xe436eb7c, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_RGB24               = new Guid(0xe436eb7d, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_RGB32               = new Guid(0xe436eb7e, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_Overlay             = new Guid(0xe436eb7f, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);

        public static Guid MEDIASUBTYPE_MPEGPacket          = new Guid(0xe436eb80, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_MPEG1Payload        = new Guid(0xe436eb81, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_MPEG1SystemStream   = new Guid(0xe436eb82, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_MPEG1System         = new Guid(0xe436eb84, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_MPEG1VideoCD        = new Guid(0xe436eb85, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_MPEG1Video          = new Guid(0xe436eb86, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_MPEG1Audio          = new Guid(0xe436eb87, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_Avi                 = new Guid(0xe436eb88, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_QTMovie             = new Guid(0xe436eb89, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_PCMAudio_Obsolete   = new Guid(0xe436eb8a, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_WAVE                = new Guid(0xe436eb8b, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_AU                  = new Guid(0xe436eb8c, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);
        public static Guid MEDIASUBTYPE_AIFF                = new Guid(0xe436eb8d, 0x524f, 0x11ce, 0x9f, 0x53, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70);

        public static Guid MEDIASUBTYPE_YUY2 = new Guid(0x32595559, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_UYVY = new Guid(0x59565955, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_YVU9 = new Guid(0x39555659, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_IYUV = new Guid(0x56555949, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_I420 = new Guid(0x30323449, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
    
        public static Guid MEDIASUBTYPE_PCM                 = new Guid(0x00000001, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71);
        public static Guid MEDIASUBTYPE_MPEG1AudioPayload   = new Guid(0x00000050, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71);
        public static Guid MEDIASUBTYPE_CLPL = new Guid(0x4C504C43, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_YUYV = new Guid(0x56595559, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_Y411 = new Guid(0x31313459, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_Y41P = new Guid(0x50313459, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_YVYU = new Guid(0x55595659, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_Y211 = new Guid(0x31313259, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_YV12 = new Guid(0x32315659, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_CLJR = new Guid(0x524a4c43, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_IF09 = new Guid(0x39304649, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_CPLA = new Guid(0x414c5043, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_MJPG0 = new Guid(0x47504A4D, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_TVMJ = new Guid(0x4A4D5654, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_WAKE = new Guid(0x454B4157, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_CFCC = new Guid(0x43434643, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_IJPG = new Guid(0x47504A49, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_Plum = new Guid(0x6D756C50, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_DVCS = new Guid(0x53435644, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_DVSD = new Guid(0x44535644, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_MDVF = new Guid(0x4656444D, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_ARGB1555 = new Guid(0x297c55af, 0xe209, 0x4cb3, 0xb7, 0x57, 0xc7, 0x6d, 0x6b, 0x9c, 0x88, 0xa8);
        public static Guid MEDIASUBTYPE_ARGB4444 = new Guid(0x6e6415e6, 0x5c24, 0x425f, 0x93, 0xcd, 0x80, 0x10, 0x2b, 0x3d, 0x1c, 0xca);
        public static Guid MEDIASUBTYPE_ARGB32 = new Guid(0x773c9ac0, 0x3274, 0x11d0, 0xb7, 0x24, 0x0, 0xaa, 0x0, 0x6c, 0x1a, 0x1);
        public static Guid MEDIASUBTYPE_AYUV = new Guid(0x56555941, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_AI44 = new Guid(0x34344941, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_IA44 = new Guid(0x34344149, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_Asf = new Guid(0x3db80f90, 0x9412, 0x11d1, 0xad, 0xed, 0x0, 0x0, 0xf8, 0x75, 0x4b, 0x99);
        public static Guid MEDIASUBTYPE_QTRpza = new Guid(0x617a7072, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_QTSmc = new Guid(0x20636d73, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_QTRle = new Guid(0x20656c72, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_QTJpeg = new Guid(0x6765706a, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_dvsd = new Guid(0x64737664, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_dvhd = new Guid(0x64687664, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_dvsl = new Guid(0x6c737664, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_Line21_BytePair = new Guid(0x6e8d4a22, 0x310c, 0x11d0, 0xb7, 0x9a, 0x0, 0xaa, 0x0, 0x37, 0x67, 0xa7);
        public static Guid MEDIASUBTYPE_Line21_GOPPacket = new Guid(0x6e8d4a23, 0x310c, 0x11d0, 0xb7, 0x9a, 0x0, 0xaa, 0x0, 0x37, 0x67, 0xa7);
        public static Guid MEDIASUBTYPE_Line21_VBIRawData = new Guid(0x6e8d4a24, 0x310c, 0x11d0, 0xb7, 0x9a, 0x0, 0xaa, 0x0, 0x37, 0x67, 0xa7);
        public static Guid MEDIASUBTYPE_TELETEXT = new Guid(0xf72a76e3, 0xeb0a, 0x11d0, 0xac, 0xe4, 0x00, 0x00, 0xc0, 0xcc, 0x16, 0xba);
        public static Guid MEDIASUBTYPE_DRM_Audio = new Guid(0x00000009, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_IEEE_FLOAT = new Guid(0x00000003, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_DOLBY_AC3_SPDIF = new Guid(0x00000092, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_RAW_SPORT = new Guid(0x00000240, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_SPDIF_TAG_241h = new Guid(0x00000241, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71);
        public static Guid MEDIASUBTYPE_DssVideo = new Guid(0xa0af4f81, 0xe163, 0x11d0, 0xba, 0xd9, 0x0, 0x60, 0x97, 0x44, 0x11, 0x1a);
        public static Guid MEDIASUBTYPE_DssAudio = new Guid(0xa0af4f82, 0xe163, 0x11d0, 0xba, 0xd9, 0x0, 0x60, 0x97, 0x44, 0x11, 0x1a);
        public static Guid MEDIASUBTYPE_VPVideo = new Guid(0x5a9b6a40, 0x1a22, 0x11d1, 0xba, 0xd9, 0x0, 0x60, 0x97, 0x44, 0x11, 0x1a);
        public static Guid MEDIASUBTYPE_VPVBI = new Guid(0x5a9b6a41, 0x1a22, 0x11d1, 0xba, 0xd9, 0x0, 0x60, 0x97, 0x44, 0x11, 0x1a);
        public static Guid MEDIASUBTYPE_UNDOC_Y422 = new Guid(0x32323459, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71);
    
        public static Guid WMMEDIASUBTYPE_mp43 = new Guid(0x3334706d, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 
        public static Guid WMMEDIASUBTYPE_MP43 = new Guid(0x3334504D, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 
        public static Guid WMMEDIASUBTYPE_mp4s = new Guid(0x7334706D, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 
        public static Guid WMMEDIASUBTYPE_MP4S = new Guid(0x5334504D, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 
        public static Guid WMMEDIASUBTYPE_WMV1 = new Guid(0x31564D57, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 
        public static Guid WMMEDIASUBTYPE_WMV2 = new Guid(0x32564D57, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 
        public static Guid WMMEDIASUBTYPE_WMV3 = new Guid(0x33564D57, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 
        public static Guid WMMEDIASUBTYPE_WMVP = new Guid(0x50564D57, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 
        public static Guid WMMEDIASUBTYPE_WMVA = new Guid(0x41564D57, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 
        public static Guid WMMEDIASUBTYPE_WVP2 = new Guid(0x32505657, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 
		public static Guid WMMEDIASUBTYPE_MSS1 = new Guid(0x3153534D, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 
		public static Guid WMMEDIASUBTYPE_MSS2 = new Guid(0x3253534D, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 
		public static Guid WMMEDIASUBTYPE_PCM = new Guid(0x00000001, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 
        public static Guid WMMEDIASUBTYPE_DRM = new Guid(0x00000009, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 
        public static Guid WMMEDIASUBTYPE_WMAudioV7 = new Guid(0x00000161, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 
        public static Guid WMMEDIASUBTYPE_WMAudioV2 = new Guid(0x00000161, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 
        public static Guid WMMEDIASUBTYPE_ACELPnet = new Guid(0x00000130, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71); 


        #endregion

        #region Format Types

        public static Guid FORMAT_None = new Guid(0x0F6417D6, 0xc318, 0x11d0, 0xa4, 0x3f, 0x00, 0xa0, 0xc9, 0x22, 0x31, 0x96);
        public static Guid FORMAT_VideoInfo = new Guid(0x05589f80, 0xc356, 0x11ce, 0xbf, 0x01, 0x00, 0xaa, 0x00, 0x55, 0x59, 0x5a);
        public static Guid FORMAT_VideoInfo2 = new Guid(0xf72a76A0, 0xeb0a, 0x11d0, 0xac, 0xe4, 0x00, 0x00, 0xc0, 0xcc, 0x16, 0xba);
        public static Guid FORMAT_WaveFormatEx = new Guid(0x05589f81, 0xc356, 0x11ce, 0xbf, 0x01, 0x00, 0xaa, 0x00, 0x55, 0x59, 0x5a);
        public static Guid FORMAT_MPEGVideo = new Guid(0x05589f82, 0xc356, 0x11ce, 0xbf, 0x01, 0x00, 0xaa, 0x00, 0x55, 0x59, 0x5a);
        public static Guid FORMAT_MPEGStreams = new Guid(0x05589f83, 0xc356, 0x11ce, 0xbf, 0x01, 0x00, 0xaa, 0x00, 0x55, 0x59, 0x5a);
        public static Guid FORMAT_DvInfo = new Guid(0x05589f84, 0xc356, 0x11ce, 0xbf, 0x01, 0x00, 0xaa, 0x00, 0x55, 0x59, 0x5a);

        
        #endregion
    }

    public class MediaTypeVideoInfo : MediaType
    {
        public MediaTypeVideoInfo(){}
        public MediaTypeVideoInfo(_AMMediaType mt) : base(mt)
        {
            VideoInfo = this.GetVideoInfo(mt.pbFormat, mt.cbFormat);
        }
        public MediaTypeVideoInfo(IntPtr pMT) : base(pMT)
        {
            //Pri1: Bug, not using pMT
            VideoInfo = this.GetVideoInfo(mt.pbFormat, mt.cbFormat);
        }
        public VIDEOINFOHEADER VideoInfo;
        public override _AMMediaType ToAMMediaType()
        {
            Update();
            return mt;
        }

        public void Update()
        {
            mt.formattype = FormatTypeEnumToGuid(FormatType);
            mt.majortype = MajorTypeEnumToGuid(MajorType);
            mt.subtype = SubTypeEnumToGuid(SubType);
            mt.bFixedSizeSamples = this.FixedSizeSamples;
            mt.bTemporalCompression = this.TemporalCompression;
            mt.lSampleSize = (uint)this.SampleSize;

            if (mt.cbFormat == 0)
            {
                // Allocate the unmanaged memory
                mt.cbFormat = (uint)Marshal.SizeOf(VideoInfo);
                mt.pbFormat = Marshal.AllocCoTaskMem((int)mt.cbFormat);
            }
            System.Runtime.InteropServices.Marshal.StructureToPtr(VideoInfo, mt.pbFormat, false);
        }
    }

    public class MediaTypeWaveFormatEx : MediaType
    {
        public MediaTypeWaveFormatEx(){}
        public MediaTypeWaveFormatEx(MSR.LST.MDShow._AMMediaType mt) : base(mt)
        {
            WaveFormatEx = this.GetWaveFormatEx(mt.pbFormat, mt.cbFormat);

        }
        public MediaTypeWaveFormatEx(IntPtr pMT) : base(pMT)
        {
            WaveFormatEx = this.GetWaveFormatEx(mt.pbFormat, mt.cbFormat);
        }
        public WAVEFORMATEX WaveFormatEx;
    }


    public enum MajorType
    {
        Video,
        Audio,
        Text,
        Midi,
        Stream,
        Interleaved,
        File,
        ScriptCommand,
        AUXLine21Data,
        VBI,
        Timecode,
        LMRT,
        UrlStream
    }

    public enum SubType
    {
        Unknown,
        WMV1,
        WMV2,
        WMV3,
        WMVP,
        WMVA,
        WVP2,
        CLPL,
        YUYV,
        IYUV,
        YVU9,
        Y411,
        Y41P,
        YUY2,
        YVYU,
        UYVY,
        Y211,
        YV12,
        CLJR,
        IF09,
        CPLA,
        MJPG0,
        TVMJ,
        WAKE,
        CFCC,
        IJPG,
        Plum,
        DVCS,
        DVSD,
        MDVF,
        RGB1,
        RGB4,
        RGB8,
        RGB565,
        RGB555,
        RGB24,
        RGB32,
        ARGB1555,
        ARGB4444,
        ARGB32,
        AYUV,
        AI44,
        IA44,
        Overlay,
        MPEGPacket,
        MPEG1Payload,
        MPEG1AudioPayload,
        MPEG1SystemStream,
        MPEG1System,
        MPEG1VideoCD,
        MPEG1Video,
        MPEG1Audio,
        Avi,
        Asf,
        QTMovie,
        QTRpza,
        QTSmc,
        QTRle,
        QTJpeg,
        PCMAudio_Obsolete,
        PCM,
        WAVE,
        AU,
        AIFF,
        dvsd,
        dvhd,
        dvsl,
        Line21_BytePair,
        Line21_GOPPacket,
        Line21_VBIRawData,
        TELETEXT,
        DRM_Audio,
        IEEE_FLOAT,
        DOLBY_AC3_SPDIF,
        RAW_SPORT,
        SPDIF_TAG_241h,
        DssVideo,
        DssAudio,
        VPVideo,
        VPVBI,
        MP43,
        mp43,
        MP4S,
        mp4s,
        MSS1,
		MSS2,
        WMPCM,
        DRM,
        WMAudioV7,
        WMAudioV2,
        ACELPnet,
        UNDOC_Y422,
        I420
    }
    public enum FormatType
    {
        None,
        VideoInfo,
        VideoInfo2,
        WaveFormatEx,
        MPEGVideo,
        MPEGStreams,
        DvInfo
    }
    public enum BitmapCompression
    {
        RGB = 0,
        RLE8 = 1,
        RLE4 = 2,
        BITFIELDS = 3,
        JPEG = 4,
        PNG = 5
    }

    /// <summary>
    /// Located in Platform SDK\include\WinDefs.h, WTypes.h, WTypes.idl
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int left;
        public int top;
        public int right;
        public int bottom;
    }
   
    /// <summary>
    /// Located in Platform SDK\include\WinDef.h, WTypes.h
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct SIZE
    {
        public int cx;
        public int cy;
    }

    /// <summary>
    /// Located in DirectX SDK\include\Amvideo.h
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct VIDEOINFOHEADER
    {
        public RECT Source;
        public RECT Target;
        public uint BitRate;
        public uint BitErrorRate;
        public UInt64 AvgTimePerFrame;
        public BITMAPINFOHEADER BitmapInfo;
    }

    /// <summary>
    /// Located in Platform SDK\include\WinGDI.h
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct BITMAPINFOHEADER
    {
        public uint Size;
        public int Width;
        public int Height;
        public ushort Planes;
        public ushort BitCount;
        public uint Compression;
        public uint SizeImage;
        public int XPelsPerMeter;
        public int YPelsPerMeter;
        public uint ClrUsed;
        public uint ClrImportant;
    }

    /// <summary>
    /// Located in Platform SDK\include\MMReg.h
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct WAVEFORMATEX
    {
        public ushort FormatTag;      
        public ushort Channels;       
        public uint   SamplesPerSec;  
        public uint   AvgBytesPerSec; 
        public ushort BlockAlign;     
        public ushort BitsPerSample;  
        public ushort Size;          
    }
   
    /// <summary>
    /// Located in DirectX SDK\include\strmif.h
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct VIDEO_STREAM_CONFIG_CAPS
    {
        public Guid guid;
        public uint VideoStandard;
        public SIZE InputSize;
        public SIZE MinCroppingSize;
        public SIZE MaxCroppingSize;
        public int  CropGranularityX;
        public int  CropGranularityY;
        public int  CropAlignX;
        public int  CropAlignY;
        public SIZE MinOutputSize;
        public SIZE MaxOutputSize;
        public int  OutputGranularityX;
        public int  OutputGranularityY;
        public int  StretchTapsX;
        public int  StretchTapsY;
        public int  ShrinkTapsX;
        public int  ShrinkTapsY;
        public long MinFrameInterval;
        public long MaxFrameInterval;
        public int  MinBitsPerSecond;
        public int  MaxBitsPerSecond;
    }
   
    /// <summary>
    /// Located in DirectX SDK\include\strmif.h
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct AUDIO_STREAM_CONFIG_CAPS
    {
        public Guid guid;
        public uint MinimumChannels;
        public uint MaximumChannels;
        public uint ChannelsGranularity;
        public uint MinimumBitsPerSample;
        public uint MaximumBitsPerSample;
        public uint BitsPerSampleGranularity;
        public uint MinimumSampleFrequency;
        public uint MaximumSampleFrequency;
        public uint SampleFrequencyGranularity;
    }
}
