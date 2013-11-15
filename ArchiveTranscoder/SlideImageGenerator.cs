using System;
using System.Diagnostics;
using System.Text;
using System.Collections;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Drawing;
using System.Drawing.Imaging;
using MSR.LST.RTDocuments;

/// NOTE: The namespace "Microsoft.Office.Interop.PowerPoint" is not available until
/// Office XP Primary Interop Assemblies have been installed on the dev machine. See:
/// http://www.microsoft.com/downloads/details.aspx?FamilyId=C41BD61E-3060-4F71-A6B4-01FEBA508E52&displaylang=en
/// Interop assemblies will be generated automatically by Visual Studio at design time, 
/// but these are considered by Microsoft to be "unofficial", and should not be used. 

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using UW.ClassroomPresenter.Model.Presentation;
using System.Configuration;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Find out which slide decks are referenced across all segments in the job and build a directory of 
	/// jpg images for each.
	/// </summary>
	public class SlideImageGenerator
	{
        #region Singleton Instance

        private static SlideImageGenerator instance = null;
        public static SlideImageGenerator Instance {
            get { return instance; }
        }
        /// <summary>
        /// The SlideImageGenerator can be shared by a set of modules.  
        /// This allows us to avoid duplicating work of building slide images.
        /// Use this static method to get the instance, then run Process if
        /// the job has changed. 
        /// </summary>
        /// <param name="job"></param>
        /// <param name="progressTracker"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        public static SlideImageGenerator GetInstance(ArchiveTranscoderJob job, 
                                            ProgressTracker progressTracker, 
                                            LogMgr log) {
            if (instance == null) {
                instance = new SlideImageGenerator(job, progressTracker, log);
            }
            else {
                //Reset all properties
                instance.Job = job;
                instance.TheProgressTracker = progressTracker;
                instance.TheLogMgr = log;
                instance.TheRtDocuments = null;
                instance.TheSlideMessages = null;
            }

            return instance;
        }

        #endregion Singleton Instance

		#region Members

		private ProgressTracker progressTracker;
		private ArchiveTranscoderJob job;
		LogMgr log;
		private Hashtable outputDirs;
		private bool stopNow;
		private Hashtable rtDocuments;
        private Hashtable slideMessages;
        private PowerPoint.Application pptApp = null;
        private bool pptAlreadyOpen = false;
        private int rightShift = 0;
        private bool useNativeJpegExport = false;
        private bool usePPTImageExport = false;

		#endregion Members

		#region Properties
		
		/// <summary>
		/// Key is the guid, Value is the path
		/// </summary>
		public Hashtable OutputDirs
		{
			get { return outputDirs; }
		}

        public ArchiveTranscoderJob Job {
            get { return this.job; }
            set { this.job = value; }
        }

        public ProgressTracker TheProgressTracker {
            set { this.progressTracker = value; }
        }

        public LogMgr TheLogMgr {
            set { this.log = value; }
        }

        /// <summary>
        /// If RtDocuments are used, this property needs to be set before processing.
        /// </summary>
        public Hashtable TheRtDocuments {
            set { this.rtDocuments = value; }
        }

        /// <summary>
        /// If legacy CP slide messages are used, this property needs to be set before processing.
        /// </summary>
        public Hashtable TheSlideMessages {
            set { this.slideMessages = value; }
        }

		#endregion Properties

		#region Ctor/Dtor

		private SlideImageGenerator(ArchiveTranscoderJob job, ProgressTracker progressTracker, 
            LogMgr log)
		{
			stopNow = false;
			this.job = job;
			this.log = log;
			this.progressTracker = progressTracker;
			this.rtDocuments = null;
            this.slideMessages = null;
			outputDirs = new Hashtable();

            //This is to work around an issue with CP3 which can cause elements on slides to be shifted a little bit to the right,
            //causing archived ink to be misaligned.  To work around, we shift the slide images to approximate what CP3 did.
            rightShift = 0;
            if (ConfigurationManager.AppSettings["RightShiftPPTImage"] != null) {
                int s;
                if (int.TryParse(ConfigurationManager.AppSettings["RightShiftPPTImage"], out s)) {
                    if (s > 0) {
                        Console.WriteLine("Right shifting images by: " + s.ToString());
                        rightShift = s;
                    }
                }
            }

            // By default we use CP3 to export PPT slide images because we get the image alignment that matches 
            // ink, and get the same appearance that was seen in the classroom.  However the following two options 
            // allow us to revert to exporting images from the PowerPoint application instead.

            //While we can get better quality from PPT images using export to wmf, some unusual characters will not 
            //appear correctly (eg. the lambda character converts as a '?').  However if we export from PPT to JPEG,
            //somewhat more of the characters are correct.

            /// Set this option to make PPT export directly to JPG
            useNativeJpegExport = false;
            if (ConfigurationManager.AppSettings["PPT2JpegExport"] != null) {
                bool b;
                if (bool.TryParse(ConfigurationManager.AppSettings["PPT2JpegExport"], out b)) {
                    Console.WriteLine("Using native PPT to Jpeg export.");
                    useNativeJpegExport = b;                   
                }
            }

            /// Set this option to get the WMF export from PPT.
            usePPTImageExport = false;
            if (ConfigurationManager.AppSettings["PPTImageExport"] != null) {
                bool b;
                if (bool.TryParse(ConfigurationManager.AppSettings["PPTImageExport"], out b)) {
                    Console.WriteLine("Using PPT for image export.");
                    usePPTImageExport = b;
                }
            }

        }

		~SlideImageGenerator()
		{
			foreach (String s in outputDirs.Values)
			{
				Directory.Delete(s,true);
			}
		}

		#endregion Ctor/Dtor

		#region Public Methods

		/// <summary>
		/// Stop the Process method if it is running
		/// </summary>
		public void Stop()
		{
			stopNow = true;
		}

		/// <summary>
		/// Build a list of slide decks referenced in the job, and generate a directory of slide images for each deck.
		/// </summary>
		public void Process()
		{
			progressTracker.CustomMessage = "Processing Slides";
			progressTracker.CurrentValue=0;

			//build a list of decks from the job spec
			Hashtable workList = new Hashtable();
			log.WriteLine("Generating Slide Images.");
			foreach (ArchiveTranscoderJobSegment segment in job.Segment)
			{
				if (segment.SlideDecks != null)
				{
					foreach (ArchiveTranscoderJobSlideDeck deck in segment.SlideDecks)
					{
						Guid deckGuid = new Guid(deck.DeckGuid);
                        if (this.outputDirs.ContainsKey(deckGuid)) { 
                            //Deck was previously converted.
                            continue;
                        }

						if (!workList.ContainsKey(deckGuid))
						{
							if (deck.Path == null)
							{
								//unmatched deck
								log.WriteLine("Warning: Specified deck has not been matched to a file: " + deck.Title );
								log.ErrorLevel = 5;
								continue;
							}
							FileInfo fi = new FileInfo(deck.Path);
							if (fi.Exists)
							{
								workList.Add(deckGuid,fi);
							}
							else
							{
                                DirectoryInfo di = new DirectoryInfo(deck.Path);
                                if (di.Exists) {
                                    workList.Add(deckGuid, di);
                                }
                                else {
                                    log.WriteLine("Warning: Specified deck does not exist: " + fi.FullName);
                                    log.ErrorLevel = 5;
                                }
							}
						}
						else if (deck.Path != ((FileInfo)workList[deckGuid]).FullName)
						{
							log.WriteLine("Warning: The same deck was specified with different paths: " + 
								deck.Path + " " + ((FileInfo)workList[deckGuid]).FullName);
							log.ErrorLevel = 3;
						}
					}
				}
			}
			log.WriteLine("Found " + workList.Count.ToString() + " deck(s) in the job.");

			//Since it is impractical to find out how many slides are in each deck here, we will consider each deck to 
			//be worth 100 in progressTracker.
			int deckCount = workList.Count;
			if (rtDocuments != null)
			{
				deckCount += rtDocuments.Count;
			}
            if (slideMessages != null)
            { 
                //For now we ignore these since there are assumed to be just a few at most.
            }

			progressTracker.EndValue = deckCount * 100;

			//Build image directory for each deck
			if (workList.Count > 0)
			{
				int currentDeck = 0;
				foreach (Guid g in workList.Keys)
				{
                    if (workList[g] is FileInfo) {
                        if (((FileInfo)workList[g]).FullName.ToLower().IndexOf(".ppt", 1) > 0) {
                            if (Utility.CheckPptIsInstalled()) {
                                if (this.useNativeJpegExport || this.usePPTImageExport) {
                                    // Use Powerpoint directly to do the image export
                                    processPpt(g, (FileInfo)workList[g]);
                                }
                                else { 
                                    // Use CP3 to open the PPT file and do the export
                                    processPptWithCp3(g, (FileInfo)workList[g]);
                                }
                            }
                            else {
                                log.WriteLine("Warning: A PowerPoint deck was specified, but PowerPoint does not appear to be " +
                                    "installed on this system.  This deck will be ignored: " +
                                    ((FileInfo)workList[g]).FullName + "  To resolve this problem specify a CSD deck in the job, " +
                                    "or run Archive Transcoder on a system with PowerPoint installed.");
                                log.ErrorLevel = 5;
                            }
                            
                        }
                        else if (((FileInfo)workList[g]).FullName.ToLower().IndexOf(".csd", 1) > 0) {
                            processCsd(g, (FileInfo)workList[g]);
                        }
                        else if (((FileInfo)workList[g]).FullName.ToLower().IndexOf(".cp3", 1) > 0) {
                            processCp3(g, (FileInfo)workList[g]);
                        }
                        else {
                            log.WriteLine("Warning: unsupported deck file type: " + ((FileInfo)workList[g]).FullName +
                                "  This deck will be ignored.");
                            log.ErrorLevel = 5;
                        }
                    }
                    else if (workList[g] is DirectoryInfo) {
                        processDir(g, (DirectoryInfo)workList[g]);
                    }
					currentDeck++;
				}
			}

            if (pptApp != null) {
                if (!pptAlreadyOpen) {
                    try {
                        pptApp.Quit();
                    }
                    catch { }
                }
                pptApp = null;
                GC.Collect();
            }


            if ((!stopNow) && (this.rtDocuments != null))
            {
                foreach (Guid g in rtDocuments.Keys)
                {
                    processRTDocument((RTDocument)rtDocuments[g]);
                }
            }

            if ((!stopNow) && (this.slideMessages != null))
            {
                foreach (String s in slideMessages.Keys)
                {
                    processSlideMessage((WorkSpace.SlideMessage)slideMessages[s]);
                }
            }


		}

		#endregion Public Methods

		#region Private Methods

        private void processSlideMessage(WorkSpace.SlideMessage sm)
        {
            if (stopNow)
                return;
            if (sm == null)
            {
                //progressTracker.CurrentValue += 100; //We are not participating in the count for now.
                return;
            }

            String tempDirName;
            if (outputDirs.ContainsKey(sm.DocRef.GUID))
            {
                tempDirName = (String)outputDirs[sm.DocRef.GUID];
            }
            else
            {
                tempDirName = Utility.GetTempDir();
                this.outputDirs.Add(sm.DocRef.GUID, tempDirName);
                Directory.CreateDirectory(tempDirName);
            }

            Image img = sm.Slide.GetImage(SlideViewer.SlideMode.Default);

            if (img != null)
            {
                //Quick poll result comes out with a black background!  We make a default background and overlay the bitmap on top.
                Image img2 = new Bitmap(img.Width, img.Height);
                SolidBrush bgBrush = new SolidBrush(Color.PapayaWhip);
                Graphics g = Graphics.FromImage(img2);
                g.FillRectangle(bgBrush, new Rectangle(0, 0, img.Width, img.Height));
                g.DrawImage(img, new Point(0, 0));

                String jpgfile = Path.Combine(tempDirName, "slide" + (sm.PageRef.Index + 1).ToString() + ".jpg");

                if (!System.IO.File.Exists(jpgfile))
                {
                    img2.Save(jpgfile, ImageFormat.Jpeg);
                }
                else
                { 
                    Debug.WriteLine("SlideImageGenerator: skipping SlideMessage: duplicate image.");
                }
            }
            else
            {
                Debug.WriteLine("SlideImageGenerator: skipping SlideMessage: null image.");
            }

        }

		private void processRTDocument(RTDocument rtd)
		{
			if (stopNow)
				return;
			if (rtd==null)
			{
				progressTracker.CurrentValue += 100;
				return;
			}

			if (outputDirs.ContainsKey(rtd.Identifier))
			{
				log.WriteLine("Warning: A duplicate RTDocument with identifier: " + rtd.Identifier + " was found in " +
					"the presentation data.  This RTDocument will be ignored.");
				log.ErrorLevel=5;
				progressTracker.CurrentValue += 100;
				return;
			}

			if (rtd.Organization.TableOfContents.Length <= 0)
			{
				progressTracker.CurrentValue += 100;
				return;
			}

			String tempDirName = Utility.GetTempDir();
			Directory.CreateDirectory(tempDirName);
			int i = 1;
			int pagecount = rtd.Organization.TableOfContents.Length;
			int ptStart = progressTracker.CurrentValue;
			foreach (TOCNode n in rtd.Organization.TableOfContents)
			{
				Page p = rtd.Resources.Pages[n.ResourceIdentifier];
				if ((p != null ) && (p.Image != null))
				{
					String jpgfile = Path.Combine(tempDirName,"slide" + i.ToString() + ".jpg");
					
					if (!repairAspectRatio(p.Image,jpgfile))
					{
						if (p.Image.Size.Width > 1600)
						{
							//PRI2: how do we make the images look better, but still be 'small'?
							Image img2 = new Bitmap(p.Image,new Size(960,720));
							img2.Save(jpgfile,ImageFormat.Jpeg);
							img2.Dispose();
						}
						else
						{			
							p.Image.Save(jpgfile,ImageFormat.Jpeg);					
						}
					}
				}
				else
				{
					Debug.WriteLine("SlideImageGenerator: skipping rtdoc page: null page or null image.");
				}
				progressTracker.CurrentValue = ((100*i)/pagecount) + ptStart;
				i++;
			}
			this.outputDirs.Add(rtd.Identifier,tempDirName);
		}

		/// <summary>
		/// Screen shots produce images with non-standard aspect ratio.  If this image has a non-standard 
		/// aspect ratio, paste it into an image that has a standard 1.33 ratio, save the result, and 
		/// return true.  Otherwise return false.
		/// </summary>
		/// <param name="img"></param>
		/// <param name="jpgFile"></param>
		/// <returns></returns>
		private bool repairAspectRatio(Image img, String jpgFile)
		{
			double ar = ((double)img.Width)/((double)img.Height);
			if (ar < 1.32)
			{
				//in this case the image is not wide enough.
				double w = ((double)img.Height) * 1.333;
				int newWidth = (int)Math.Round(w);
				Bitmap b = new Bitmap(newWidth,img.Height);
				Graphics g = Graphics.FromImage(b);
				SolidBrush whiteBrush = new SolidBrush(Color.White);
				//make a white background
				g.FillRectangle(whiteBrush,0,0,b.Width,b.Height);
				//paste the original image on the top of the new bitmap.
				g.DrawImage(img,0,0,img.Width,img.Height);

				if (b.Width > 1600)
				{
					Image img2 = new Bitmap(b,new Size(960,720));
					img2.Save(jpgFile,ImageFormat.Jpeg);
					img2.Dispose();
				}
				else
				{
					b.Save(jpgFile,ImageFormat.Jpeg);
				}

				whiteBrush.Dispose();
				b.Dispose();
				g.Dispose();
				return true;
			}
			else if (ar > 1.34)
			{
				//in this case the image is not tall enough.
				double h = ((double)img.Width)/1.333;
				int newHeight = (int)Math.Round(h);
				Bitmap b = new Bitmap(img.Width,newHeight);
				Graphics g = Graphics.FromImage(b);
				SolidBrush whiteBrush = new SolidBrush(Color.White);
				//make a white background
				g.FillRectangle(whiteBrush,0,0,b.Width,b.Height);
				//paste the original image on the top of the new bitmap.
				g.DrawImage(img,0,0,img.Width,img.Height);

				if (b.Width > 1600)
				{
					Image img2 = new Bitmap(b,new Size(960,720));
					img2.Save(jpgFile,ImageFormat.Jpeg);
					img2.Dispose();
				}
				else
				{
					b.Save(jpgFile,ImageFormat.Jpeg);
				}

				whiteBrush.Dispose();
				b.Dispose();
				g.Dispose();
				return true;
			}
			return false;
		}

        /// <summary>
        /// The deck provided is already in the form of a directory of images
        /// </summary>
        /// <param name="g"></param>
        /// <param name="directoryInfo"></param>
        private void processDir(Guid g, DirectoryInfo sourceDir) {
            String tempDirName = Utility.GetTempDir();
            Directory.CreateDirectory(tempDirName);
            //Here we just assume the files are all correctly named..
            foreach (FileInfo fi in sourceDir.GetFiles()) {
                fi.CopyTo(Path.Combine(tempDirName,fi.Name));
            }
            this.outputDirs.Add(g, tempDirName);
        }


		/// <summary>
		/// Open a PPT deck and create a directory of jpg images, one for each slide in the deck.
		/// </summary>
		/// <param name="g"></param>
		/// <param name="file"></param>
		private void processPpt(Guid g, FileInfo file)
		{
			if (stopNow)
				return;

            if (pptApp == null) {
                try {
                    pptApp = new PowerPoint.Application();
                    // Check if PPT was already running.
                    pptAlreadyOpen = (pptApp.Visible == MsoTriState.msoTrue);
                }
                catch (Exception e) {
                    log.WriteLine("Failed to create PowerPoint application instance. " + e.ToString());
                    log.ErrorLevel = 6;
                    pptApp = null;
                    return;
                }
            }

            try {
                PowerPoint._Presentation ppPres = pptApp.Presentations.Open(file.FullName, MsoTriState.msoTrue,
                    MsoTriState.msoFalse, MsoTriState.msoFalse);

                PowerPoint.Slides slides = ppPres.Slides;

                String tempDirName = Utility.GetTempDir();
                Directory.CreateDirectory(tempDirName);
                int slideCount = slides.Count;
                int ptStart = progressTracker.CurrentValue;
                for (int i = 1; i <= slides.Count; i++) {
                    if (stopNow)
                        break;
                    //The image looks much better and is reasonably small if we save to wmf format first, then 
                    // convert to jpg.  I guess the .Net jpeg encoder is better than the PPT encoder??
                    // Sometimes we may want to use the jpeg export with certain decks because we can get better
                    // fidelity with some less common symbols.

                    PowerPoint._Slide ppSlide = (PowerPoint._Slide)slides._Index(i);
                    String jpgfile = Path.Combine(tempDirName, "slide" + i.ToString() + ".jpg");
                    Image img = null;
                    string tmpfile = null;
                    if (useNativeJpegExport) { 
                        tmpfile = Path.Combine(tempDirName, "slidetmp" + i.ToString() + ".JPG");
                        ppSlide.Export(tmpfile, "JPG", 0, 0);
                        img = Image.FromFile(tmpfile);
                    }
                    else {
                        tmpfile = Path.Combine(tempDirName, "slide" + i.ToString() + ".WMF");
                        ppSlide.Export(tmpfile, "WMF", 0, 0);
                        img = Image.FromFile(tmpfile);
                    }
                    try {
                        //The right shift is used to work around a CP3 issue: Some slide decks are shifted to the right, so we also
                        // need to do this to keep ink from being mis-aligned.  As of 2013 this should no longer be needed when
                        // using a current version of CP3.
                        if (rightShift > 0) {
                            Image shiftedImage;
                            RightImageShift(img, rightShift, out shiftedImage);
                            shiftedImage.Save(jpgfile, ImageFormat.Jpeg);
                            shiftedImage.Dispose();
                        }
                        else {
                            img.Save(jpgfile, ImageFormat.Jpeg);
                        }

                    }
                    catch (Exception e) {
                        log.WriteLine("Failed to save image for slide " + jpgfile + " exception: " + e.ToString());
                        log.ErrorLevel = 6;
                    }
                    img.Dispose();
                    System.IO.File.Delete(tmpfile);

                    progressTracker.CurrentValue = ((i * 100) / slideCount) + ptStart;
                }

                ppPres.Close();
                ppPres = null;

                this.outputDirs.Add(g, tempDirName);
            }
            catch (Exception e) {
                log.WriteLine(e.ToString());
                log.ErrorLevel = 6;
            }
		}

        /// <summary>
        /// Open the file with CP3, then use CP3's export function to create slide images.
        /// </summary>
        /// <param name="g"></param>
        /// <param name="fileInfo"></param>
        private void processPptWithCp3(Guid g, FileInfo fileInfo) {
            if (stopNow)
                return;

            //TODO: can we do anything with progress tracker here?
            try {
                DeckModel deck;
                deck = UW.ClassroomPresenter.Decks.PPTDeckIO.OpenPPT(fileInfo);
                String tempDirName = Utility.GetTempDir();
                DefaultDeckTraversalModel traversal = new SlideDeckTraversalModel(Guid.NewGuid(), deck);

                if (stopNow)
                    return;

                List<string> imageNames = UW.ClassroomPresenter.Decks.PPTDeckIO.ExportDeck(traversal, 
                    tempDirName, ImageFormat.Jpeg);

                fixSlideImageNames(tempDirName);

                this.outputDirs.Add(g, tempDirName);
            }
            catch (Exception ex)             
            {
                log.WriteLine(ex.ToString());
                log.ErrorLevel = 6;
            }
         }

        /// <summary>
        /// If slide images are created using a different naming convention, rename them to match 
        /// the ArchiveTranscoder convention, eg.: "slide1.jpg"
        /// </summary>
        /// <param name="tempDirName"></param>
        private void fixSlideImageNames(string tempDirName) {
            DirectoryInfo di = new DirectoryInfo(tempDirName);
            FileInfo[] files = di.GetFiles("*.jpg");
            for (int i = 0; i < files.Length; i++) {
                //Remove the file extension.
                string basename = files[i].Name.Substring(0, files[i].Name.LastIndexOf("."));
                //Match a string ending in a non-digit followed by one or more digits, then end of string.
                Regex re = new Regex(@"^.*\D(\d+)$");
                Match m = re.Match(basename);
                if (m.Success) {
                    int slidenum;
                    if (int.TryParse(m.Groups[1].ToString(), out slidenum)) {
                        string dest = Path.Combine(tempDirName, "slide" + slidenum.ToString() + ".jpg");
                        files[i].MoveTo(dest);
                    }
                    else { 
                        throw new ApplicationException("Unexpected slide image number.  Failed to parse slide number from the name: " + basename);                      
                    }
                }
                else {
                    throw new ApplicationException("Unexpected slide image file name.  Failed to parse slide number from the name: " + basename);
                }
            }

        }


        /// <summary>
        /// This was added to correct an alignment problem between cp3 and ppt image export.
        /// </summary>
        /// <param name="sourceImage"></param>
        /// <param name="shiftBy"></param>
        /// <param name="destImage"></param>
        private void RightImageShift(Image sourceImage, int shiftBy, out Image destImage) {
            Image tmpImage = new Bitmap(960, 720);
            Graphics g = Graphics.FromImage(tmpImage);
            g.DrawImage(sourceImage, new Rectangle(new Point(0, 0), new Size(960, 720)));
            g.DrawImage(sourceImage, new Rectangle(new Point(shiftBy, 0), new Size(960, 720)));
            g.Dispose();
            destImage = tmpImage;
        }

        /// <summary>
        /// Generate slide images from a CP3 file
        /// </summary>
        /// <param name="g"></param>
        /// <param name="fileInfo"></param>
        private void processCp3(Guid g, FileInfo fileInfo) {
            if (stopNow)
                return;

            String tempDirName = Utility.GetTempDir();

            try {
                DeckModel deck = UW.ClassroomPresenter.Decks.PPTDeckIO.OpenCP3(fileInfo);
                SlideDeckTraversalModel dtm = new SlideDeckTraversalModel(Guid.NewGuid(), deck);
                if (stopNow)
                    return;
                UW.ClassroomPresenter.Decks.PPTDeckIO.ExportDeck(dtm, tempDirName, ImageFormat.Jpeg);
                dtm.Dispose();
            }
            catch (Exception e) { 
                log.WriteLine("Failed to open CP3 file: " + fileInfo.FullName +
                    "  Exception: " + e.ToString());
                log.ErrorLevel = 5;
                return;
            }

            //Make conformant file names
            tempDirName = translateCP3ImageNames(tempDirName);

            this.outputDirs.Add(g, tempDirName);
        }

        private string translateCP3ImageNames(string inputDir) {
            string outputdir = Utility.GetTempDir();
            Directory.CreateDirectory(outputdir);

            DirectoryInfo inputDi = new DirectoryInfo(inputDir);
            foreach (FileInfo fi in inputDi.GetFiles()) {
                try {
                    string basename = Path.GetFileNameWithoutExtension(fi.Name);
                    string index = basename.Substring(basename.LastIndexOf("_") + 1);
                    int i = Int32.Parse(index);
                    string outname = Path.Combine(outputdir, "slide" + i.ToString() + ".jpg");
                    fi.CopyTo(outname);
                }
                catch (Exception e) {
                    log.WriteLine("Failed to export slide.  Exception: " + e.ToString());
                    log.ErrorLevel = 5;
                    continue;
                }
            }

            try {
                Directory.Delete(inputDir, true);
            }
            catch (Exception e) {
                log.WriteLine("Failed to delete temp directory.  Exception: " + e.ToString());
                log.ErrorLevel = 5;     
            }

            return outputdir;
        }

		/// <summary>
		/// Open a CSD and create a directory of jpg images, one per slide.
		/// </summary>
		/// <param name="g"></param>
		/// <param name="file"></param>
		private void processCsd(Guid g, FileInfo file)
		{
			if (stopNow)
				return;

			String tempDirName = Utility.GetTempDir();

			SlideViewer.CSDDocument document = null;
			BinaryFormatter binaryFormatter = new BinaryFormatter();
			FileStream stream = null;
			try
			{
				stream = new FileStream(file.FullName,FileMode.Open,FileAccess.Read,FileShare.ReadWrite);
				document = (SlideViewer.CSDDocument)binaryFormatter.Deserialize(stream);
			}
			catch (Exception e)
			{
				Debug.WriteLine("Failed to read: " + file.FullName + "; exception text: " + e.ToString());
				log.WriteLine("Error: Failed to read file: " + file.FullName + 
                    ".  Possibly the file was produced with an unsupported version of Classroom Presenter.  Exception text: " +
                    e.ToString());
				log.ErrorLevel = 5;
			}
			finally
			{
				stream.Close();
			}

            if (document == null)
                return;


			if (document.DocumentType!="Presentation")
			{
				log.WriteLine("Error: Document type is not Presentation: " + file.FullName);
				log.ErrorLevel = 3;
				progressTracker.CurrentValue += 100;
				return;
			}

			Directory.CreateDirectory(tempDirName);

			int pageNum = 1;
			int ptStart = progressTracker.CurrentValue;
			foreach (SlideViewer.SlideDocument page in document.Pages)
			{
				if (stopNow)
					break;
				// SlideMode images may be any or all of Student, Instructor, Shared, Default..  
				// I think we prefer Default, but will also take Shared.
				if (page.slideImages.Count > 0)
				{
					object o = null;
					foreach (String s in page.slideImages.Keys)
					{
						if (s=="Default")
						{
							o = page.slideImages[s];
							break;
						}
						else if ((s=="Shared") && (o==null))
						{
							o = page.slideImages[s];
						}
					}
					//The image may be SerializableImage or just Image, the former for EMF.
					if (o!=null)
					{
						Image img = null;
						if (o is SlideViewer.SerializableImage)
						{
							SlideViewer.SerializableImage si = (SlideViewer.SerializableImage)o;
							img = si.image;
						}
						else if (o is System.Drawing.Image)
						{
							img = (Image)o;
						}
						else
						{
							log.WriteLine("Error: Unrecognized image type in deck: " + file.FullName );
							log.ErrorLevel = 5;
						}
						if (img != null)
						{
							String outfile = Path.Combine(tempDirName,"slide" + pageNum.ToString() + ".jpg");
							//The images in EMF format are rather huge.. scale them down.
							if (img.Size.Width > 1600)
							{
								Image img2 = new Bitmap(img,new Size(960,720));
								img2.Save(outfile,ImageFormat.Jpeg);
								//PRI2: the scaled images tend to look like crap. Fix?
								// try converting to wmf, then to jpg?
								//this.SaveWithQuality(outfile,img2,100); //<-- no help
								img2.Dispose();
								log.WriteLine ("Warning: Slide " + pageNum.ToString() + " was scaled. " +
									"For better results, use a CSD that was created with WMF format, or specify the " +
									"original PPT file.");
								log.ErrorLevel = 3;
							}
							else
							{
								img.Save(outfile,System.Drawing.Imaging.ImageFormat.Jpeg);
							}
							img.Dispose();
						}
					}
				}
				progressTracker.CurrentValue = ((pageNum*100)/document.Pages.Length) + ptStart;
				pageNum++;
			}
			log.WriteLine("Processed " + (pageNum-1).ToString() + " pages for " + file.FullName);
			outputDirs.Add(g,tempDirName);
		}

		#region Unused

		private void SaveWithQuality(String outfile, Image img, int quality)
		{
			ImageCodecInfo myImageCodecInfo;
			System.Drawing.Imaging.Encoder myEncoder;
			EncoderParameter myEncoderParameter;
			EncoderParameters myEncoderParameters;
			myImageCodecInfo = GetEncoderInfo("image/jpeg");
			myEncoder = System.Drawing.Imaging.Encoder.Quality;
			myEncoderParameters = new EncoderParameters(1);
			myEncoderParameter = new EncoderParameter(myEncoder, quality);
			myEncoderParameters.Param[0] = myEncoderParameter;
			img.Save(outfile, myImageCodecInfo, myEncoderParameters);
		}

		private static ImageCodecInfo GetEncoderInfo(String mimeType)
		{
			int j;
			ImageCodecInfo[] encoders;
			encoders = ImageCodecInfo.GetImageEncoders();
			for(j = 0; j < encoders.Length; ++j)
			{
				if(encoders[j].MimeType == mimeType)
					return encoders[j];
			}
			return null;
		}

		#endregion Unused

		#endregion Private Methods
	}
}
