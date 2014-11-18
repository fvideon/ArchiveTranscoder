using System;
using System.Collections.Generic;
using System.Text;
using MSR.LST.Net.Rtp;
using MSR.LST;
using MSR.LST.MDShow;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.IO;
using ArchiveRTNav;
using ImageFilter;

namespace ArchiveTranscoder {
    /// <summary>
    /// From a video stream, produce a set of stills and metadata to allow using the stills as a presentation.
    /// </summary>
    public class PresentationFromVideoMgr {
        #region Constants

        /// <summary>
        /// Default for the minimum time between frame grabs (in msec)
        /// </summary>
        private const int FRAMEGRAB_INTERVAL_MS = 500;

        /// <summary>
        /// Default difference threshold for imagemagick compare.
        /// </summary>
        private const double DIFFERENCE_THRESHOLD = 270;

        /// <summary>
        /// If true, scale HD images down by reducing dimensions by half.  If false keep the source image size.
        /// </summary>
        private const bool SCALE = false;

        /// <summary>
        /// If true, the source is pillarboxed widescreen format from which we should attempt to restore a 
        /// 4x3 frame dimension.
        /// </summary>
        private const bool PILLARBOXED = true;

        #endregion Constants

        #region Private Members

        private ArchiveTranscoderJobSegmentVideoDescriptor videoDescriptor;
        private long start;
        private long end;
        private long offset;
        private long previousSegmentEnd;
        private LogMgr errorLog;
        private ProgressTracker progressTracker;
        private bool stopNow;
        private int frameGrabIntervalMs = FRAMEGRAB_INTERVAL_MS;
        private long lastFramegrab = 0;
        private string tempdir;
        private List<PresentationMgr.DataItem> metadata;
        private StreamMgr videoStream;
        private Guid deckGuid;
        private List<double> differenceMetrics = null;
        private double differenceThreshold = DIFFERENCE_THRESHOLD;

        #endregion Private Members

        #region Properties

        /// <summary>
        /// The minimum time in msec between frame grabs.
        /// </summary>
        public int FrameGrabInterval {
            get { return this.frameGrabIntervalMs; }
            set { this.frameGrabIntervalMs = value; }
        }

        public List<double> FrameGrabDifferenceMetrics {
            get { return this.differenceMetrics; }
        }

        /// <summary>
        /// Files that differ by more than this threshold will be retained during the
        /// initial processing.
        /// </summary>
        public double DifferenceThreshold {
            get { return this.differenceThreshold;  }
            set { this.differenceThreshold = value; }
        }

        /// <summary>
        /// The path to the temp directory that will contain the image files after 
        /// the Process method has completed.
        /// </summary>
        public string TempDir {
            get { return this.tempdir; }
        }

        public Guid DeckGuid {
            get { return this.deckGuid; }
        }

        #endregion Properties

        #region Public Methods

        public PresentationFromVideoMgr(ArchiveTranscoderJobSegmentVideoDescriptor videoDescriptor,
            long start, long end, long offset, LogMgr errorLog, long previousSegmentEnd, ProgressTracker progressTracker) {
            this.videoDescriptor = videoDescriptor;
            this.start = start;
            this.end = end;
            this.offset = offset;
            this.errorLog = errorLog;
            this.previousSegmentEnd = previousSegmentEnd;
            this.progressTracker = progressTracker;
            this.tempdir = null;
            this.videoStream = null;
            this.deckGuid = Guid.Empty;
        }

        /// <summary>
        /// Capture stills from the video stream.  Create a temporary directory and save the images there.
        /// Compile a list of DataItem objects to indicate when the slide transitions should take place.
        /// </summary>
        /// <returns></returns>
        public string Process() {
            ImageFilter.ImageFilter imgFilter = null;
            try {
                imgFilter = new ImageFilter.ImageFilter();
                imgFilter.DifferenceThreshold = this.differenceThreshold;
            }
            catch {
                this.errorLog.Append("Video capture images in the presentation will not be filtered probably " +
                    "because ImageMagick is not available in the configuration.\r\n");
            }

            this.differenceMetrics = new List<double>();
            metadata = new List<PresentationMgr.DataItem>();
            RTUpdate rtu = new RTUpdate();
            this.deckGuid = Guid.NewGuid();
            rtu.DeckGuid = this.deckGuid;
            rtu.SlideSize = 1.0;
            rtu.DeckType = (Int32)DeckTypeEnum.Presentation;

            this.videoStream = new StreamMgr(videoDescriptor.VideoCname, videoDescriptor.VideoName, new  DateTime(this.start), new DateTime(this.end), false, PayloadType.dynamicVideo);
            this.videoStream.ToRawWMFile(this.progressTracker); 
            MediaTypeVideoInfo mtvi = videoStream.GetUncompressedVideoMediaType();

            this.tempdir = Utility.GetTempDir();
            Directory.CreateDirectory(this.tempdir);
            string filebase = "slide";
            string extent = ".jpg";
            int fileindex = 1;

            BufferChunk bc;
            long time;
            bool newStream;
            string previousFile = null;
            this.stopNow = false;
            while ((!stopNow) && (videoStream.GetNextSample(out bc, out time, out newStream))) {
                if ((time - lastFramegrab) >= (long)(this.frameGrabIntervalMs * Constants.TicksPerMs)) {
                    DateTime dt = new DateTime(time);
                    Debug.WriteLine("time=" + dt.ToString() + ";length=" + bc.Length.ToString());
                    lastFramegrab = time;
                    string filepath = Path.Combine(tempdir, filebase + fileindex.ToString() + extent);
                    PixelFormat pixelFormat = subtypeToPixelFormat(mtvi.SubType); 
                    Bitmap bm = new Bitmap(mtvi.VideoInfo.BitmapInfo.Width, mtvi.VideoInfo.BitmapInfo.Height, pixelFormat);
                    BitmapData bd = bm.LockBits(new Rectangle(0, 0, mtvi.VideoInfo.BitmapInfo.Width, mtvi.VideoInfo.BitmapInfo.Height), ImageLockMode.ReadWrite, pixelFormat);
                    Marshal.Copy(bc.Buffer, 0, bd.Scan0, bc.Length);
                    bm.UnlockBits(bd);
                    bm.RotateFlip(RotateFlipType.RotateNoneFlipY);
                    if ((SCALE) && (mtvi.VideoInfo.BitmapInfo.Width >= 1280)) {
                        int w = mtvi.VideoInfo.BitmapInfo.Width / 2;
                        int h = mtvi.VideoInfo.BitmapInfo.Height / 2;
                        Bitmap scaled = new Bitmap(bm, new Size(w,h));
                        scaled.Save(filepath, ImageFormat.Jpeg);
                    }
                    else {
                        bm.Save(filepath, ImageFormat.Jpeg);
                    }
                    if (PILLARBOXED) {
                        bm = UnPillerBox(bm);
                        bm.Save(filepath, ImageFormat.Jpeg);
                    }
                    if (imgFilter != null) {
                        string filterMsg;
                        bool filterError;
                        double metric;
                        bool differ = imgFilter.ImagesDiffer(filepath, previousFile, out filterMsg, out filterError, out metric);
                        if (filterError) {
                            //this.errorLog.Append(filterMsg);
                            Console.WriteLine(filterMsg);
                        }
                        if (!differ) continue;
                        this.differenceMetrics.Add(metric);
                        Console.WriteLine("Framegrab slide index: " + fileindex.ToString() +
                            "; difference: " + metric.ToString() + "; time: " + dt.ToString());

                    }
                    rtu.SlideIndex = fileindex-1; // This is a zero-based index
                    metadata.Add(new PresentationMgr.DataItem(time-this.offset, PresentationMgr.CopyRTUpdate(rtu)));
                    fileindex++;
                    previousFile = filepath;
                }            
            }
            return null;
        }


        /// <summary>
        /// specific to sorting files like this "slide1.jpg", "slide2.jpg", "slide10.jpg", etc
        /// </summary>
        private class PMPComparer : IComparer<string> {
            public int Compare(string x, string y) {
                int xi; int yi;
                if (Int32.TryParse(Path.GetFileNameWithoutExtension(x).Substring(5), out xi) &&
                    Int32.TryParse(Path.GetFileNameWithoutExtension(y).Substring(5), out yi)) {
                    if (xi == yi) return 0;
                    else if (xi < yi) return -1;
                    else return 1;
                }
                return 0;
            }
        }

        /// <summary>
        /// Filter the contents of the existing tempdir using the specified
        /// difference threshold.  Rewrite tempdir, metadata and differenceMetrics.
        /// </summary>
        /// <param name="threshold"></param>
        internal void Refilter(double threshold) {
            //Make an imageFilter instance and configure with the new threshold.
            ImageFilter.ImageFilter imgFilter = null;
            try {
                imgFilter = new ImageFilter.ImageFilter();
                imgFilter.DifferenceThreshold = threshold;
            }
            catch {
                return;
            }
            Console.WriteLine("Refiltering framegrabs using threshold: " + threshold.ToString());

            List<double> newDifferenceMetrics = new List<double>();
            List<PresentationMgr.DataItem> newMetadata = new List<PresentationMgr.DataItem>();
            RTUpdate rtu = new RTUpdate();
            this.deckGuid = Guid.NewGuid();
            rtu.DeckGuid = this.deckGuid;
            rtu.SlideSize = 1.0;
            rtu.DeckType = (Int32)DeckTypeEnum.Presentation;
            string filebase = "slide";
            string extent = ".jpg";
            int fileindex = 1;

            try {
                string newDir = Utility.GetTempDir();
                Directory.CreateDirectory(newDir);
                string[] files = Directory.GetFiles(this.tempdir, "*.jpg");
                Array.Sort(files, new PMPComparer());
                string previous = null;
                for (int i = 0; i < files.Length; i++) {
                    string file = files[i];
                    string msg;
                    bool error = false;
                    double metric;

                    if (imgFilter.ImagesDiffer(previous, file, out msg, out error, out metric)) {
                        string newFilename = filebase + fileindex.ToString() + extent;
                        File.Copy(file, Path.Combine(newDir, newFilename));
                        rtu.SlideIndex = fileindex - 1; // This is a zero-based index
                        newMetadata.Add(new PresentationMgr.DataItem(metadata[i].Timestamp, PresentationMgr.CopyRTUpdate(rtu)));
                        newDifferenceMetrics.Add(metric);
                        DateTime t = new DateTime(metadata[i].Timestamp);
                        Console.WriteLine("Refilter keeping old index: " + (i+1).ToString() + 
                            "; new index: " + fileindex.ToString() +
                            "; difference: " + metric.ToString() + "; time: " + t.ToString());
                        fileindex++;
                        previous = file;
                    }
                }
                this.metadata = newMetadata;
                this.differenceMetrics = newDifferenceMetrics;
                Directory.Delete(this.tempdir, true);
                this.tempdir = newDir;
            }
            catch (Exception) {
                return;
            }

        }
 
        public void Stop() {
            this.stopNow = true;
            if (this.videoStream != null) {
                this.videoStream.Stop();
            }
        }

        public string Print() {
            StringBuilder sb = new StringBuilder();
            sb.Append(GetSlideTitles());
            foreach (PresentationMgr.DataItem di in metadata) {
                sb.Append(di.Print());
            }
            return sb.ToString();
        } 

        #endregion Public Methods

        #region Private Methods

        private string GetSlideTitles() {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < metadata.Count; i++) {
                sb.Append("<Title DeckGuid=\"" + this.deckGuid.ToString() + "\" Index=\"" + i.ToString() +
                    "\" Text=\"Video Capture\"/>  ");

            }
            return sb.ToString();
        }

        private PixelFormat subtypeToPixelFormat(SubType subType) {
            switch (subType) {
                case SubType.RGB24: {
                    return PixelFormat.Format24bppRgb;
                }
                case SubType.RGB555: {
                    return PixelFormat.Format16bppRgb555;
                }
                case SubType.RGB32: {
                    return PixelFormat.Format32bppRgb;
                }
                case SubType.RGB565: {
                    return PixelFormat.Format16bppRgb565;
                }
                default: {
                    return PixelFormat.Format24bppRgb;
                }
            }
        }

        /// <summary>
        /// Assume the input bitmap is pillarboxed.  Return a bitmap with 4x3 aspect ratio
        /// and pillarboxing removed.
        /// </summary>
        /// <param name="bm"></param>
        /// <returns></returns>
        internal Bitmap UnPillerBox(Bitmap inImg) {
            if ((float)inImg.Width / (float)inImg.Height <= 4f / 3f) {
                //Nothing to do
                return inImg;
            }

            int midw = inImg.Width / 2;  // Half the width of the original
            int targetw = inImg.Height * 4 / 3;  // Image width we want in the output
            int lbb = midw - (targetw / 2);  // The x coordinate of the crop bounding box
            Rectangle cropR = new Rectangle(lbb, 0, targetw, inImg.Height);
            Bitmap outImg = new Bitmap(cropR.Width, cropR.Height, inImg.PixelFormat);
            using (Graphics g = Graphics.FromImage(outImg)) {
                g.DrawImage(inImg, new Rectangle(0, 0, cropR.Width, cropR.Height), cropR, GraphicsUnit.Pixel);
            }

            return outImg;
        }

        #endregion Private Methods
               
        #region Filtering Framegrabs

        //Note: In practice we are likely to get about 1.5 times this number of images in the output,
        // so we set it lower than we really want.
        private static int MAX_FRAMEGRABS_PER_JOB = 350;

        /// <summary>
        /// This should be called after all segments are processed.  Take the results
        /// and if there are frame grabs, figure out if they should be filtered further.
        /// </summary>
        /// <param name="segments"></param>
        internal static void FilterSegments(List<WMSegment> segments) {
            int fgCount = getFrameGrabCount(segments);
            if (fgCount > MAX_FRAMEGRABS_PER_JOB) {
                List<double> diffMetrics = getDifferenceMetrics(segments);
                diffMetrics.Sort();
                if (diffMetrics.Count > MAX_FRAMEGRABS_PER_JOB) {
                    double newThreshold = diffMetrics[diffMetrics.Count - MAX_FRAMEGRABS_PER_JOB - 1];
                    foreach (WMSegment s in segments) {
                        if (s.FrameGrabMgr != null) {
                            s.FrameGrabMgr.Refilter(newThreshold);
                        }
                    }
                    fgCount = getFrameGrabCount(segments);
                }
            }

            ////TODO: Not sure it's even possible to filter toc entries here.  I think the html5 translation code
            //// can do this, and might need new webviewer code too.
            //if (fgCount > MAX_TOC_ENTRIES_PER_JOB) {
            //    //refresh with revised thresholds
            //    List<double> diffMetrics = getDifferenceMetrics(segments);
            //    diffMetrics.Sort();
            //    double newTOCThreshold = diffMetrics[diffMetrics.Count - MAX_TOC_ENTRIES_PER_JOB - 1];
            //    foreach (WMSegment s in segments) {
            //        //TODO: here we just want to attach the threshold to each FrameGrabMgr so that when
            //        // titles are printed we skip the ones below the threshold.
            //    }
            //}
        }

        /// <summary>
        /// Aggregate a list of difference metrics across all segments
        /// </summary>
        /// <param name="segments"></param>
        /// <returns></returns>
        private static List<double> getDifferenceMetrics(List<WMSegment> segments) {
            List<double> result = new List<double>();
            foreach (WMSegment s in segments) {
                if (s.FrameGrabDifferenceMetrics != null) {
                    result.AddRange(s.FrameGrabDifferenceMetrics);
                }
            }
            return result;
        }

        /// <summary>
        /// Peek inside all framegrab directories to get the total file count.
        /// </summary>
        /// <param name="segments"></param>
        /// <returns></returns>
        private static int getFrameGrabCount(List<WMSegment> segments) {
            int cnt = 0;
            foreach (WMSegment s in segments) {
                if (s.FrameGrabDirectory != null) {
                    if (Directory.Exists(s.FrameGrabDirectory)) {
                        DirectoryInfo di = new DirectoryInfo(s.FrameGrabDirectory);
                        cnt += di.GetFiles("*.jpg").Length;
                    }
                }
            }
            return cnt;
        }
    #endregion Filtering

        #region Unit Test

        public static Bitmap TestUnPillarBox(Bitmap bm) {
            PresentationFromVideoMgr instance = new PresentationFromVideoMgr(null, 0, 0, 0, null, 0, null);
            return instance.UnPillerBox(bm);
        }
        
        #endregion Unit Test
    }
        
}
