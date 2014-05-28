using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.Serialization.Formatters.Binary;
using MSR.LST;
using MSR.LST.MDShow;
using MSR.LST.Net.Rtp;
using MSR.LST.RTDocuments;
using System.Diagnostics;
using WorkSpace;
using System.IO;
using System.Collections;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace ArchiveTranscoder
{
    /// <summary>
    /// Generate a sequence of video samples from a presentation source.  Create an instance of SlideImageMgr to 
    /// maintain a map of the slides and decks.  Feed presentation data to the SlideImageMgr, and query it
    /// for the current frame at the interval appropriate to enforce the requested frame rate.
    /// </summary>
    class SlideStreamMgr:IDisposable {

        #region Members

        private SlideImageGenerator imageGenerator = null;
        private SlideImageMgr slideImageMgr;
        private ArchiveTranscoderJobSegment segment;
        private ArchiveTranscoderJob job;
        private LogMgr logMgr;
        private MediaTypeVideoInfo uncompressedMT;
        private bool cancel;
        private DateTime start;
        private DateTime end;
        private int[] streamIDs;
        private PayloadType payload;
        private PresenterWireFormatType format;
        private DBStreamPlayer[] streamPlayers;
        private string cname;
        private bool initialized;
        private bool pptInstalled;
        private double lookBehindDuration;
        private long ticksBetweenFrames;
        private long nextFrameTime;
        private int outputWidth = 320;
        private int outputHeight = 240;
        
        #endregion Members

        #region Properties
        /// <summary>
        /// The MediaType describing uncompressed samples we deliver
        /// </summary>
        public MediaTypeVideoInfo UncompressedMediaType
        {
            get 
            {
                return uncompressedMT;
            }
        }

        #endregion Properties

        #region Ctor

        public SlideStreamMgr(ArchiveTranscoderJob job, ArchiveTranscoderJobSegment segment, 
            LogMgr logMgr, double fps, int width, int height)
        {
            this.job = job;
            this.segment = segment;
            this.logMgr = logMgr;
            
            if (width > 0 && height > 0)
            {
                this.outputWidth = width;
                this.outputHeight = height;
            }

            this.ticksBetweenFrames = (long)((double)Constants.TicksPerSec / fps);

            uncompressedMT = getUncompressedMT(this.outputWidth, this.outputHeight, fps);
            cancel = false;
            initialized = false;
            pptInstalled = Utility.CheckPptIsInstalled();
            
            if ((!DateTime.TryParse(segment.StartTime, out start)) ||
                (!DateTime.TryParse(segment.EndTime,out end)))
            {
                throw(new System.Exception("Failed to parse start/end time"));
            }

            this.nextFrameTime = start.Ticks;

            format = Utility.StringToPresenterWireFormatType(segment.PresentationDescriptor.PresentationFormat);
            payload = Utility.formatToPayload(format);
            cname = segment.PresentationDescriptor.PresentationCname;

            slideImageMgr = new SlideImageMgr(format, this.outputWidth, this.outputHeight);

            //Get the start time for the entire conference and use that to get streams.
            long confStart = DatabaseUtility.GetConferenceStartTime(payload, cname, start.Ticks, end.Ticks);
            if (confStart <= 0)
            {
                logMgr.WriteLine("Warning: No conference exists in the database that matches this presentation: " + cname + 
                    " with PresentationFormat " + format.ToString());
                logMgr.ErrorLevel = 7;
                confStart = start.Ticks;
            }

            //Get the relevant stream_id's and create DBStreamPlayers for each.
            streamIDs = DatabaseUtility.GetStreams(payload,segment.PresentationDescriptor.PresentationCname, null, confStart, end.Ticks);
            DateTime sdt = new DateTime(confStart);
            Debug.WriteLine("***Conference start: " + sdt.ToString() + " end: " + end.ToString());
            if ((streamIDs == null) || (streamIDs.Length==0))
            {
                Debug.WriteLine("No presentation data found.");
                logMgr.WriteLine("Warning: No presentation data was found for the given time range for " +
                    cname + " with PresentationFormat " + format.ToString());
                logMgr.ErrorLevel = 7;
                streamPlayers = null;
                return;
            }

            streamPlayers = new DBStreamPlayer[streamIDs.Length];
            for (int i = 0; i < streamIDs.Length; i++)
            {
                streamPlayers[i] = new DBStreamPlayer(streamIDs[i], confStart, end.Ticks, payload);
            }

            lookBehindDuration = 0;
            if (streamPlayers[0].Start < start)
            {
                lookBehindDuration = ((TimeSpan)(start - streamPlayers[0].Start)).TotalSeconds;
            }

        }

        #endregion Ctor

        #region IDisposable Members

        public void Dispose()
        {
            if (slideImageMgr != null)
            {
                slideImageMgr.Dispose();
                slideImageMgr = null;
            }
        }

        #endregion IDisposable Members

        #region Public Methods

        /// <summary>
        /// Do all the initial time consuming preparation including building temp directories from
        /// any decks specified with the job, and scanning data prior to start time to attempt to determine
        /// the presentation state.
        /// We require that Init complete before GetNextSample.
        /// </summary>
        public bool Init(ProgressTracker progressTracker)
        {
            if ((streamPlayers == null) || (streamPlayers.Length == 0)) {
                // No data is not considered an error.
                return true;
            }

            //preserve the end value to be restored when we're done.
            int oldEnd = progressTracker.EndValue;

            //build directories from any decks specified in the job
            imageGenerator = SlideImageGenerator.GetInstance(this.job, progressTracker, logMgr);

            //Refresh image export size and job in case they have changed.
            imageGenerator.SetImageExportSize(false, this.outputWidth, this.outputHeight);
            imageGenerator.Job = this.job;

            //Run process to build or refresh deck images.
            imageGenerator.Process();

            //Tell SlideImageMgr about the decks.
            slideImageMgr.SetSlideDirs(imageGenerator.OutputDirs);

            //scan data preceeding start time to establish all initial state
            progressTracker.CurrentValue = 0;
            progressTracker.EndValue = (int)lookBehindDuration;
            progressTracker.CustomMessage = "Initializing Presentation Data";
            BufferChunk bc;
            long timestamp;
            while (!cancel)
            {
                long t; int index;
                if (getNextStreamPlayerFrameTime(out t, out index))
                {
                    if (t < start.Ticks)
                    {
                        if (streamPlayers[index].GetNextFrame(out bc, out timestamp))
                        {
                            slideImageMgr.ProcessFrame(bc);
                            progressTracker.CurrentValue = (int)(((TimeSpan)(new DateTime(timestamp) - streamPlayers[0].Start)).TotalSeconds);
                        }
                        else
                        {
                            break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }
                else
                {
                    break;
                }
            }

            if (!cancel)
                initialized = true;

            progressTracker.CustomMessage = "";
            progressTracker.EndValue = oldEnd;
            return true;
        }

        /// <summary>
        /// Return the next uncompressed sample in the stream.
        /// Only valid after Init has completed.
        /// </summary>
        public bool GetNextSample(out BufferChunk sample, out long time, out bool newstream)
        { 
            sample = null;
            time = 0;
            newstream = false;

            if (!initialized)
                return false;

            if ((streamPlayers == null) || (streamPlayers.Length == 0))
                return false;

            if (nextFrameTime > end.Ticks)
                return false;

            long t;
            int index;

            while (!cancel)
            {
                if (!getNextStreamPlayerFrameTime(out t, out index))
                {
                    //No more data
                    t = long.MaxValue;
                }

                if (t < start.Ticks)
                {
                    //shouldn't happen
                    Debug.Assert(false);
                }
                else if (t >= this.nextFrameTime)
                {
                    //We have already read to the next frame time, so return a sample.
                    slideImageMgr.GetCurrentFrame(out sample);
                    time = nextFrameTime;
                    if (time == start.Ticks)
                        newstream = true;
                    nextFrameTime += this.ticksBetweenFrames;
                    return true;
                }
                else if (t < end.Ticks)
                {
                    //Get the next presentation data frame and feed it to SlideImageMgr, then loop again.
                    BufferChunk bc;
                    long timestamp;
                    if (streamPlayers[index].GetNextFrame(out bc, out timestamp))
                    {
                        slideImageMgr.ProcessFrame(bc);
                    }
                }
            }

            //if cancelled:
            return false;
        }


        /// <summary>
        /// Cancel an init operation in progress.
        /// </summary>
        public void Stop()
        {
            cancel = true;
            if (imageGenerator != null)
                imageGenerator.Stop();
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Return the media type representing the uncompressed frames we will deliver to the caller.  
        /// Assume RGB24.
        /// </summary>
        /// <param name="w"></param>
        /// <param name="h"></param>
        /// <returns></returns>
        private MediaTypeVideoInfo getUncompressedMT(int w, int h, double fps)
        {
            MediaTypeVideoInfo mt = new MediaTypeVideoInfo();
            mt.FixedSizeSamples = true;
            mt.TemporalCompression = false;
            mt.SampleSize = h * w * 3;
            mt.MajorType = MajorType.Video;
            mt.SubType = SubType.RGB24;
            mt.FormatType = FormatType.VideoInfo;
            mt.VideoInfo.Source = new RECT();
            mt.VideoInfo.Source.left = 0;
            mt.VideoInfo.Source.top = 0;
            mt.VideoInfo.Source.bottom = h;
            mt.VideoInfo.Source.right = w;
            mt.VideoInfo.Target = new RECT();
            mt.VideoInfo.Target.left = 0;
            mt.VideoInfo.Target.top = 0;
            mt.VideoInfo.Target.bottom = h;
            mt.VideoInfo.Target.right = w;
            mt.VideoInfo.AvgTimePerFrame = (ulong)Constants.TicksPerSec/(ulong)fps;
            mt.VideoInfo.BitErrorRate = 0;
            mt.VideoInfo.BitRate = (uint)((double)h * (double)w * 3d * 8d * fps);
            mt.VideoInfo.BitmapInfo.Height = h;
            mt.VideoInfo.BitmapInfo.Width = w;
            mt.VideoInfo.BitmapInfo.SizeImage = (uint)(h * w * 3);
            mt.VideoInfo.BitmapInfo.Planes = 1;
            mt.VideoInfo.BitmapInfo.BitCount = 24;
            mt.VideoInfo.BitmapInfo.ClrImportant = 0;
            mt.VideoInfo.BitmapInfo.ClrUsed = 0;
            mt.VideoInfo.BitmapInfo.Compression = 0;
            mt.VideoInfo.BitmapInfo.XPelsPerMeter = 0;
            mt.VideoInfo.BitmapInfo.YPelsPerMeter = 0;
            mt.VideoInfo.BitmapInfo.Size = 40;
            mt.Update();
            return mt;
        }


        /// <summary>
        /// Iterate through the set of DBStreamPlayers to find the next frame time.
        /// We have to do this because a presentation may consist of multiple concurrent streams.
        /// (in the case of RTDocuments at least).  Return false if there are no more frames.
        /// </summary>
        private bool getNextStreamPlayerFrameTime(out long nextTime, out int nextIndex)
        {
            nextTime = long.MaxValue;
            nextIndex = 0;

            if ((streamPlayers == null) || (streamPlayers.Length == 0))
            {
                return false;
            }

            int minIndex = -1;
            long minTime = long.MaxValue;
            for (int i = 0; i < streamPlayers.Length; i++)
            {
                if (streamPlayers[i].GetNextFrameTime(out nextTime))
                {
                    if (nextTime < minTime)
                    {
                        minTime = nextTime;
                        minIndex = i;
                    }
                }
            }
            if (minIndex == -1)
            {
                return false;
            }

            nextIndex = minIndex;
            nextTime = minTime;
            return true;
        }

        #endregion Private Methods

    }
}
