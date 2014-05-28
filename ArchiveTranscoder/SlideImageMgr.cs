using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Collections;
using MSR.LST;
using MSR.LST.RTDocuments;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;


namespace ArchiveTranscoder
{
    /// <summary>
    /// Maintain collections of SlideImage objects.  Accept and demux Slide, Ink and Nav messages as well as deck references
    /// to keep the SlideImages up to date.  Maintain a pointer to the current SlideImage, and other global state, and return 
    /// the frame from the current SlideImage upon request.
    /// </summary>
    /// Note similarities between code here and in PresentationMgr..
    class SlideImageMgr:IDisposable
    {
        #region Members

        private Hashtable slideImages;          //key formed from deck/slide ids, value is a SlideImage.
        private SlideReference currentSlide;    //Contains global state: current slide ID/size/background color.
        private PresenterWireFormatType format; //Indicates the wire format of the presentation source
        private Hashtable orgToResource;        //key is the organization ID, value is the resource ID.  For RTDocuments scenarios.
        private SlideImage nullSlide;           //For the purpose of generating a blank image if we need one
        private CP3Manager.CP3Manager cp3Mgr;   //For translating a CP3 message stream.
        private int exportHeight = 0;
        private int exportWidth = 0;

        #endregion Members

        #region Construct

        public SlideImageMgr(PresenterWireFormatType format, int width, int height)
        {
            this.format = format;
            currentSlide = new SlideReference();
            this.exportWidth = width;
            this.exportHeight = height;
            
            //Set the default background colors
            if ((format == PresenterWireFormatType.CPCapability) ||
                (format == PresenterWireFormatType.CPNav))
            {
                currentSlide.SetBGColor(Color.PapayaWhip);
            }
            else
            {
                currentSlide.SetBGColor(Color.White);
            }

            if (format == PresenterWireFormatType.CP3) {
                cp3Mgr = new CP3Manager.CP3Manager();
            }

            slideImages = new Hashtable();
            orgToResource = new Hashtable();
            nullSlide = null;
        }

        #endregion Construct

        #region IDisposable Members

        public void Dispose()
        {
            if (slideImages != null)
            {
                foreach (SlideImage si in slideImages.Values)
                {
                    si.Dispose();
                }
                slideImages = null;
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// This takes the output of SlideImageGenerator.  Key is deck Guid, value is the path to the
        /// directory of images.  If there are slide directories available, call this before passing any frames.
        /// </summary>
        public void SetSlideDirs(Hashtable slideDirs)
        {
            if (slideDirs == null)
                return;

            //For each directory find qualifying image files and use them to prepopulate
            //SlideImages.
            foreach (Guid dirGuid in slideDirs.Keys)
            {
                if (Directory.Exists((String)slideDirs[dirGuid]))
                {
                    foreach (String filePath in Directory.GetFiles((String)slideDirs[dirGuid],"*.jpg"))
                    {
                        String filename = Path.GetFileNameWithoutExtension(filePath);
                        if (filename.StartsWith("slide"))
                        { 
                            int slideNum;
                            if (Int32.TryParse(filename.Substring(5), out slideNum))
                            {
                                if (slideNum > 0)
                                {
                                    int slideIndex = slideNum - 1; //Index is zero-based.
                                    String key = SlideReference.MakeKey(dirGuid, slideIndex);
                                    if (!slideImages.ContainsKey(key))
                                    {
                                        slideImages.Add(key, new SlideImage());
                                    }
                                    ((SlideImage)slideImages[key]).SetFilePath(filePath);
                                }
                            }
                        }
                    }
                }
            }
        }


        /// <summary>
        /// Accept a raw presentation frame.  Deserialize and Demux to handler for the particular presentation format.
        /// </summary>
        /// <param name="bc"></param>
        public void ProcessFrame(BufferChunk frame)
        { 
            BinaryFormatter bf = new BinaryFormatter();
            Object rtobj = null;
            try
            {
                MemoryStream ms = new MemoryStream((byte[])frame);
                rtobj = bf.Deserialize(ms);
            }
            catch (Exception e)
            {
                Debug.WriteLine("frameToSlide: exception deserializing message. size=" + frame.Length.ToString() +
                    " exception=" + e.ToString());
                return;
            }

            switch (format)
            {
                case PresenterWireFormatType.CPNav:
                    {
                        acceptCPNavFrame(rtobj);
                        break;
                    }
                case PresenterWireFormatType.CPCapability:
                    {
                        acceptCPCapabilityFrame(rtobj);
                        break;
                    }
                case PresenterWireFormatType.RTDocument:
                    {
                        acceptRTDocFrame(rtobj);
                        break;
                    }
                case PresenterWireFormatType.CP3: {
                        acceptCP3Frame(rtobj);
                        break;
                }
                default:
                    break;
            }
            
        }


        /// <summary>
        /// Return the raw bitmap representing the current presentation state
        /// </summary>
        /// <param name="frame"></param>
        public void GetCurrentFrame(out BufferChunk frame)
        {
            frame = null;

            if ((currentSlide.IsSet) && (slideImages.ContainsKey(currentSlide.GetStringHashCode())))
            {
                frame = ((SlideImage)slideImages[currentSlide.GetStringHashCode()]).GenerateFrame(currentSlide.Size, currentSlide.BGColor, this.exportWidth, this.exportHeight);
            }
            else
            {
                //If we don't have a current slide, return a blank frame.
                if (nullSlide == null)
                {
                    nullSlide = new SlideImage();
                }

                frame = nullSlide.GenerateFrame(1.0,Color.PapayaWhip, this.exportWidth, this.exportHeight);
            }
            
        }

        #endregion Public Methods

        #region Private Methods

        #region CPNav

        /// <summary>
        /// Frames sent by Classroom Presenter 2.0 in Native stand-alone mode
        /// </summary>
        /// <param name="rtobj"></param>
        private void acceptCPNavFrame(object rtobj)
        {
            if (rtobj is PresenterNav.CPPageUpdate)
            {
                /// Presenter sends these once per second. These indicate current slide index, deck index
                /// and deck type.  Just use this to update the current slide pointer
                PresenterNav.CPPageUpdate cppu = (PresenterNav.CPPageUpdate)rtobj;
                currentSlide.SetCPReference(cppu.DocRef.GUID, cppu.PageRef.Index);
                //CPPageUpdate has a deck type too, but we're not using it for now.
                //DumpPageUpdate((PresenterNav.CPPageUpdate)rtobj);
            }
            else if (rtobj is PresenterNav.CPScrollLayer)
            {
                ///Presenter sends these once per second, and during a scroll operation.
                ///Set the scroll position for the slide in question.
                PresenterNav.CPScrollLayer cpsl = (PresenterNav.CPScrollLayer)rtobj;
                String key = SlideReference.MakeKey(cpsl.DocRef.GUID, cpsl.PageRef.Index);
                if (!slideImages.ContainsKey(key))
                {
                    slideImages.Add(key, new SlideImage());
                }
                ((SlideImage)slideImages[key]).SetScrollPos(cpsl.ScrollPosition, cpsl.ScrollExtent);

                //DumpScrollLayer((PresenterNav.CPScrollLayer)rtobj);
            }
            else if (rtobj is PresenterNav.CPDeckCollection)
            {
                //Presenter sends this once per second.  The only thing we want to get from 
                //this is the slide size (former ScreenConfiguration).  This is a 'global' setting, not a per-slide setting.
                PresenterNav.CPDeckCollection cpdc = (PresenterNav.CPDeckCollection)rtobj;
                currentSlide.SetSize(cpdc.ViewPort.SlideSize);
                
                //DumpDeckCollection((PresenterNav.CPDeckCollection)rtobj);
                //The CPDeckCollection also has slide titles, but we don't care about them here.
            }
            else if (rtobj is WorkSpace.BeaconPacket) //the original beacon
            {
                //Presenter sends this once per second.  The only thing we want to get from 
                //this is the background color.  This is another global presentation setting.
                WorkSpace.BeaconPacket bp = (WorkSpace.BeaconPacket)rtobj;
                currentSlide.SetBGColor(bp.BGColor);
                //DumpBeacon((BeaconPacket)rtobj);
            }
            else if (rtobj is PresenterNav.CPDrawStroke) //add a stroke
            {
                PresenterNav.CPDrawStroke cpds = (PresenterNav.CPDrawStroke)rtobj;
                String key = SlideReference.MakeKey(cpds.DocRef.GUID, cpds.PageRef.Index);
                if (!slideImages.ContainsKey(key))
                {
                    slideImages.Add(key, new SlideImage());
                }
                ((SlideImage)slideImages[key]).AddInk(cpds.Stroke.Ink, cpds.Stroke.Guid);
            }
            else if (rtobj is PresenterNav.CPDeleteStroke) //delete one stroke
            {
                PresenterNav.CPDeleteStroke cpds = (PresenterNav.CPDeleteStroke)rtobj;
                String key = SlideReference.MakeKey(cpds.DocRef.GUID, cpds.PageRef.Index);
                if (!slideImages.ContainsKey(key))
                {
                    slideImages.Add(key, new SlideImage());
                }
                ((SlideImage)slideImages[key]).RemoveInk(cpds.Guid);
            }
            else if (rtobj is PresenterNav.CPEraseLayer) //clear all strokes from one page
            {
                PresenterNav.CPEraseLayer cpel = (PresenterNav.CPEraseLayer)rtobj;
                String key = SlideReference.MakeKey(cpel.DocRef.GUID, cpel.PageRef.Index);
                if (!slideImages.ContainsKey(key))
                {
                    slideImages.Add(key, new SlideImage());
                }
                ((SlideImage)slideImages[key]).RemoveAllInk();

            }
            else if (rtobj is PresenterNav.CPEraseAllLayers) //clear all strokes from a deck
            {
                PresenterNav.CPEraseAllLayers cpeal = (PresenterNav.CPEraseAllLayers)rtobj;
                if (cpeal.DocRef.GUID != Guid.Empty)
                {
                    foreach (String key in slideImages.Keys)
                    {
                        if (key.StartsWith(cpeal.DocRef.GUID.ToString()))
                        {
                            ((SlideImage)slideImages[key]).RemoveAllInk();
                        }
                    }
                }
            }
            else if (rtobj is WorkSpace.OverlayMessage) //Student submission received
            {
                //These are seen when Student Submissions are submitted to the instructor, not when they are displayed.
                WorkSpace.OverlayMessage om = (WorkSpace.OverlayMessage)rtobj; 
                String key = SlideReference.MakeKey(om.FeedbackDeck.GUID, om.SlideIndex.Index);
                if (!slideImages.ContainsKey(key))
                {
                    slideImages.Add(key, new SlideImage());
                }

                //The slide image comes from a presentation slide.  Here we want to store the current slide image.
                String refKey = SlideReference.MakeKey(om.PresentationDeck.GUID,om.PDeckSlideIndex.Index);
                if (slideImages.ContainsKey(refKey))
                {
                    SlideImage sourceImage = (SlideImage)slideImages[refKey];
                    bool isWhiteboard = true;
                    Image img = sourceImage.GetSlideImage(this.currentSlide.BGColor, out isWhiteboard);
                    ((SlideImage)slideImages[key]).SetImage(img,isWhiteboard);
                }

                //Add the ink from the student submission
                //OtherScribble appears to contain both student ink and any pre-existing instructor ink, but I think it
                //used to be in UserScribble.  Let's include all ink in both scribbles.
                ((SlideImage)slideImages[key]).AddInk(((SlideViewer.InkScribble)om.SlideOverlay.UserScribble).Ink);
                ((SlideImage)slideImages[key]).AddInk(((SlideViewer.InkScribble)om.SlideOverlay.OtherScribble).Ink);
            }
            else if (rtobj is WorkSpace.ScreenConfigurationMessage)
            {
                // nothing to do here.
            }
            else
            {
                Type t = rtobj.GetType();
                Debug.WriteLine("Unhandled Type:" + t.ToString());
            }
        }

        #endregion CPNav

        #region CPCapability

        /// <summary>
        /// Handle frames from Classroom Presenter when it is used as a CXP capability
        /// </summary>
        /// <param name="rtobj"></param>
        private void acceptCPCapabilityFrame(object rtobj)
        {
            // RTNodeChanged , RTStrokeAdd, and RTStrokeRemove contain a RTFrame which contains the CP message.
            if (rtobj is RTNodeChanged)
            {
                //Classroom Presenter wraps CPPageUpdate inside the RTFrame which is in turn wrapped in
                // the RTNodeChanged Extension.
                if (((RTNodeChanged)rtobj).Extension is RTFrame)
                {
                    RTFrame rtf = (RTFrame)((RTNodeChanged)rtobj).Extension;
                    if (rtf.ObjectTypeIdentifier == Constants.PageUpdateIdentifier)
                    {
                        if (rtf.Object is PresenterNav.CPPageUpdate)
                        {
                            PresenterNav.CPPageUpdate cppu = (PresenterNav.CPPageUpdate)rtf.Object;
                            currentSlide.SetCPReference(cppu.DocRef.GUID, cppu.PageRef.Index);
                        }
                    }
                }
            }
            else if (rtobj is RTStroke)
            {
                if (((RTStroke)rtobj).Extension is RTFrame)
                {
                    RTFrame rtf = (RTFrame)((RTStroke)rtobj).Extension;
                    if (rtf.ObjectTypeIdentifier == Constants.StrokeDrawIdentifier)
                    {
                        if (rtf.Object is PresenterNav.CPDrawStroke)
                        {
                            PresenterNav.CPDrawStroke cpds = (PresenterNav.CPDrawStroke)rtf.Object;
                            String key = SlideReference.MakeKey(cpds.DocRef.GUID, cpds.PageRef.Index);
                            if (!slideImages.ContainsKey(key))
                            {
                                slideImages.Add(key, new SlideImage());
                            }
                            ((SlideImage)slideImages[key]).AddInk(cpds.Stroke.Ink, cpds.Stroke.Guid);
                        }
                    }
                }
            }
            else if (rtobj is RTStrokeAdd)
            {
                if (((RTStrokeAdd)rtobj).Extension is RTFrame)
                {
                    RTFrame rtf = (RTFrame)((RTStrokeAdd)rtobj).Extension;
                    if (rtf.ObjectTypeIdentifier == Constants.StrokeDrawIdentifier)
                    {
                        if (rtf.Object is PresenterNav.CPDrawStroke)
                        {
                            PresenterNav.CPDrawStroke cpds = (PresenterNav.CPDrawStroke)rtf.Object;
                            String key = SlideReference.MakeKey(cpds.DocRef.GUID, cpds.PageRef.Index);
                            if (!slideImages.ContainsKey(key))
                            {
                                slideImages.Add(key, new SlideImage());
                            }
                            ((SlideImage)slideImages[key]).AddInk(cpds.Stroke.Ink, cpds.Stroke.Guid);
                        }
                    }
                }
            }
            else if (rtobj is RTStrokeRemove)
            {
                if (((RTStrokeRemove)rtobj).Extension is RTFrame)
                {
                    RTFrame rtf = (RTFrame)((RTStrokeRemove)rtobj).Extension;
                    if (rtf.ObjectTypeIdentifier == Constants.StrokeDeleteIdentifier)
                    {
                        if (rtf.Object is PresenterNav.CPDeleteStroke)
                        {
                            PresenterNav.CPDeleteStroke cpds = (PresenterNav.CPDeleteStroke)rtf.Object;
                            String key = SlideReference.MakeKey(cpds.DocRef.GUID, cpds.PageRef.Index);
                            if (!slideImages.ContainsKey(key))
                            {
                                slideImages.Add(key, new SlideImage());
                            }
                            ((SlideImage)slideImages[key]).RemoveInk(cpds.Guid);
                        }
                    }
                }
            }
            //Other CP messages are sent directly in RTFrame
            else if (rtobj is RTFrame)
            {
                if (((RTFrame)rtobj).ObjectTypeIdentifier == Constants.CLASSROOM_PRESENTER_MESSAGE)
                {
                    acceptCPNavFrame(((RTFrame)rtobj).Object);
                }
                else if (((RTFrame)rtobj).ObjectTypeIdentifier == Constants.RTDocEraseAllGuid)
                {
                    //This case is a bit ugly because the RTDoc message contains no info about the deck and slide.
                    // We need to apply to the current deck and slide (if any)
                    if ((currentSlide.IsSet) && (slideImages.ContainsKey(currentSlide.GetStringHashCode())))
                    {
                        ((SlideImage)slideImages[currentSlide.GetStringHashCode()]).RemoveAllInk();
                    }
                }

            }
            //Still other messages are in native format
            else if (rtobj is PresenterNav.CPMessage)
            {
                //CP still sends some messages in native format
                acceptCPNavFrame(rtobj);
            }
        }

        #endregion CPCapability

        #region RTDocument

        /// <summary>
        /// Handle frames from a native RTDocument generator such as the CXP presentation tool
        /// </summary>
        /// <param name="rtobj"></param>
        private void acceptRTDocFrame(object rtobj)
        {
            ///Notes about RTDocuments:
            ///
            /// RTDocuments have Resources and Organizations.  Resources contain Pages/Images etc while Organizations 
            /// contain the TOC/titles, etc. The TOC Nodes contain references to the resources and resource IDs.
            /// The navigation message RTNodeChanged just tells us the organization node ID, while
            /// Page and ink messages only contain the Resource ID.  RTDocument messages contain the TOC which maps pages and org nodes.
            /// PageAdds will not have an existing TocNode in the RTDocument map, but they carry their own TocNode property.
            ///
            /// For this application, we only care about having one unique page identifier.  We take the strategy of storing
            /// SlideImages under the resource Identifier, and maintaining a lookup table of Organization identifier to 
            /// resource identifier.  We use this table to resolve Organization identifiers when navigation messages are received.
            
            if (rtobj is RTDocument)
            {
                RTDocument rtd = (RTDocument)rtobj;
                //Keep the mapping of TocNode.Identifier to TocNode.ResourceIdentifier 
                foreach (TOCNode tn in rtd.Organization.TableOfContents)
                {
                    if (!orgToResource.ContainsKey(tn.Identifier))
                    {
                        orgToResource.Add(tn.Identifier, tn.ResourceIdentifier);
                    }
                    else
                    {
                        orgToResource[tn.Identifier] = tn.ResourceIdentifier;
                    }
                }

                //There is an implicit navigation to the first slide here.
                this.currentSlide.SetRTDocReference(rtd.Organization.TableOfContents[0].ResourceIdentifier);
            }
            else if (rtobj is Page)
            {
                //These are slide deck pages
                //p.Identifier is a Resource Identifier.  Store the image under that Identifier.  
                Page p = (Page)rtobj;
                if (!slideImages.ContainsKey(p.Identifier.ToString()))
                {
                    slideImages.Add(p.Identifier.ToString(), new SlideImage());
                }
                ((SlideImage)slideImages[p.Identifier.ToString()]).SetImage(p.Image,false);
            }
            else if (rtobj is RTPageAdd)
            {
                //These are dynamically added pages such as WB and screenshots
                RTPageAdd rtpa = (RTPageAdd)rtobj;
                //RTPageAdd comes with a TocNode.  Store the mapping of resource ID to TocNode.Identifier
                if (!orgToResource.ContainsKey(rtpa.TOCNode.Identifier))
                {
                    orgToResource.Add(rtpa.TOCNode.Identifier, rtpa.Page.Identifier);
                }
                else
                {
                    orgToResource[rtpa.TOCNode.Identifier] = rtpa.Page.Identifier;
                }

                //Store the page Image under the resource ID.
                if (!slideImages.ContainsKey(rtpa.Page.Identifier.ToString()))
                {
                    slideImages.Add(rtpa.Page.Identifier.ToString(), new SlideImage());
                }
                ((SlideImage)slideImages[rtpa.Page.Identifier.ToString()]).SetImage(rtpa.Page.Image,false);
            }
            else if (rtobj is RTNodeChanged)
            {
                RTNodeChanged rtnc = (RTNodeChanged)rtobj;
                //Look up the resource ID and update curent page.
                if (orgToResource.ContainsKey(rtnc.OrganizationNodeIdentifier))
                {
                    currentSlide.SetRTDocReference(((Guid)orgToResource[rtnc.OrganizationNodeIdentifier]));
                }
                else
                {
                    //Indicate slide missing by setting currentSlide reference to Guid.Empty
                    currentSlide.SetRTDocReference(Guid.Empty);
                }
            }
            else if (rtobj is RTStroke)
            {
                RTStroke rts = (RTStroke)rtobj;
                //apply the ink to the given Page Identifier.  Create a new SlideImage if necessary.
                if (!slideImages.ContainsKey(rts.PageIdentifier.ToString()))
                {
                    slideImages.Add(rts.PageIdentifier.ToString(), new SlideImage());
                }
                Microsoft.Ink.Ink ink = rts.Stroke.Ink.Clone();
                for (int i = 0; i < ink.Strokes.Count; i++)
                    ink.Strokes[i].Scale(500f / 960f, 500f / 720f);  

                ((SlideImage)slideImages[rts.PageIdentifier.ToString()]).AddInk(ink,rts.StrokeIdentifier);

                //There appears to be an implicit navigation here.  
                currentSlide.SetRTDocReference(rts.PageIdentifier);
            }
            else if (rtobj is RTStrokeRemove)
            {
                RTStrokeRemove rtsr = (RTStrokeRemove)rtobj;
                //Use the PageIdentifer to identify the page from which to remove the stroke.
                if (slideImages.ContainsKey(rtsr.PageIdentifier.ToString()))
                {
                    ((SlideImage)slideImages[rtsr.PageIdentifier.ToString()]).RemoveInk(rtsr.StrokeIdentifier);
                }
            }
            else if (rtobj is RTFrame)
            {
                RTFrame rtf = (RTFrame)rtobj;
                if (rtf.ObjectTypeIdentifier == Constants.RTDocEraseAllGuid)
                {
                    //Erase all ink on the current slide.
                    if ((currentSlide.IsSet) && (slideImages.ContainsKey(currentSlide.GetStringHashCode())))
                    {
                        ((SlideImage)slideImages[currentSlide.GetStringHashCode()]).RemoveAllInk();
                    }
                }
                else
                {
                    Debug.WriteLine("Unhandled RTFrame type.");
                }
            }
            else
            {
                Debug.WriteLine("Unhandled RT obj:" + rtobj.ToString());
            }
        }

        #endregion RTDocument

        #region CP3

        private void acceptCP3Frame(object rtobj) {
            //Feed frame to CP3Manager, then process any resulting ArchiveRTNav Frames
            List<object> archiveObjList = cp3Mgr.Accept(rtobj);
            if ((archiveObjList != null) && (archiveObjList.Count > 0)) {
                foreach (object o in archiveObjList) {
                    acceptArchiveRTNavFrame(o); 
                }
            }
        }

        private void acceptArchiveRTNavFrame(object o) {
            if (o is ArchiveRTNav.RTUpdate) { 
                //Navigation and slide size changes
                ArchiveRTNav.RTUpdate rtu = (ArchiveRTNav.RTUpdate)o;

                String key = SlideReference.MakeKey(rtu.DeckGuid, rtu.SlideIndex);
                if (!slideImages.ContainsKey(key)) {
                    slideImages.Add(key, new SlideImage());
                }

                if (!rtu.DeckAssociation.Equals(Guid.Empty)) { 
                    //Student submission or quickpoll
                    //The slide image may come from a presentation slide.  Here we want to store the current slide image.
                    String refKey = SlideReference.MakeKey(rtu.DeckAssociation, rtu.SlideAssociation);
                    if (slideImages.ContainsKey(refKey)) {

                        SlideImage sourceImage = (SlideImage)slideImages[refKey];
                        bool isWhiteboard = true;
                        Image img = sourceImage.GetSlideImage(this.currentSlide.BGColor, out isWhiteboard);
                        ((SlideImage)slideImages[key]).SetImage(img, isWhiteboard);
                    }
                }

                currentSlide.SetCPReference(rtu.DeckGuid, rtu.SlideIndex);
                currentSlide.SetSize(rtu.SlideSize);
                currentSlide.SetBGColor(rtu.BackgroundColor);
            }
            else if (o is ArchiveRTNav.RTDrawStroke) {
                ArchiveRTNav.RTDrawStroke rtds = (ArchiveRTNav.RTDrawStroke)o;
                String key = SlideReference.MakeKey(rtds.DeckGuid, rtds.SlideIndex);
                if (!slideImages.ContainsKey(key)) {
                    slideImages.Add(key, new SlideImage());
                }
                ((SlideImage)slideImages[key]).AddInk(rtds.Ink,rtds.Guid);
            }
            else if (o is ArchiveRTNav.RTDeleteStroke) {
                ArchiveRTNav.RTDeleteStroke rtds = (ArchiveRTNav.RTDeleteStroke)o;
                String key = SlideReference.MakeKey(rtds.DeckGuid, rtds.SlideIndex);
                if (!slideImages.ContainsKey(key)) {
                    slideImages.Add(key, new SlideImage());
                }
                ((SlideImage)slideImages[key]).RemoveInk(rtds.Guid);        
            }
            else if (o is ArchiveRTNav.RTTextAnnotation) {
                ArchiveRTNav.RTTextAnnotation rtta = (ArchiveRTNav.RTTextAnnotation)o;
                String key = SlideReference.MakeKey(rtta.DeckGuid, rtta.SlideIndex);
                if (!slideImages.ContainsKey(key)) {
                    slideImages.Add(key, new SlideImage());
                }
                ((SlideImage)slideImages[key]).AddTextAnnotation(rtta.Text,rtta.Origin,rtta.Guid,rtta.Font,rtta.Color,rtta.Height,rtta.Width);
            }
            else if (o is ArchiveRTNav.RTImageAnnotation) {
                ArchiveRTNav.RTImageAnnotation rtia = (ArchiveRTNav.RTImageAnnotation)o;
                String key = SlideReference.MakeKey(rtia.DeckGuid, rtia.SlideIndex);
                if (!slideImages.ContainsKey(key)) {
                    slideImages.Add(key, new SlideImage());
                }
                ((SlideImage)slideImages[key]).AddImageAnnotation(rtia.Guid, rtia.Height, rtia.Width, rtia.Origin, rtia.Img);
            }
            else if (o is ArchiveRTNav.RTDeleteTextAnnotation) {
                ArchiveRTNav.RTDeleteTextAnnotation rtdta = (ArchiveRTNav.RTDeleteTextAnnotation)o;
                String key = SlideReference.MakeKey(rtdta.DeckGuid, rtdta.SlideIndex);
                if (!slideImages.ContainsKey(key)) {
                    slideImages.Add(key, new SlideImage());
                }
                ((SlideImage)slideImages[key]).RemoveTextAnnotation(rtdta.Guid);
            }
            else if (o is ArchiveRTNav.RTDeleteAnnotation) {
                ArchiveRTNav.RTDeleteAnnotation rtda = (ArchiveRTNav.RTDeleteAnnotation)o;
                String key = SlideReference.MakeKey(rtda.DeckGuid, rtda.SlideIndex);
                if (!slideImages.ContainsKey(key)) {
                    slideImages.Add(key, new SlideImage());
                }
                ((SlideImage)slideImages[key]).RemoveAnnotation(rtda.Guid);
            }
            else if (o is ArchiveRTNav.RTEraseLayer) {
                ArchiveRTNav.RTEraseLayer rtel = (ArchiveRTNav.RTEraseLayer)o;
                String key = SlideReference.MakeKey(rtel.DeckGuid, rtel.SlideIndex);
                if (!slideImages.ContainsKey(key)) {
                    slideImages.Add(key, new SlideImage());
                }
                ((SlideImage)slideImages[key]).RemoveAllInk();
           
            }
            else if (o is ArchiveRTNav.RTEraseAllLayers) {
                ArchiveRTNav.RTEraseAllLayers rteal = (ArchiveRTNav.RTEraseAllLayers)o;
                if (rteal.DeckGuid != Guid.Empty)
                {
                    foreach (String key in slideImages.Keys)
                    {
                        if (key.StartsWith(rteal.DeckGuid.ToString()))
                        {
                            ((SlideImage)slideImages[key]).RemoveAllInk();
                        }
                    }
                }

            }
            else if (o is ArchiveRTNav.RTQuickPoll) {
                ArchiveRTNav.RTQuickPoll rtqp = (ArchiveRTNav.RTQuickPoll)o;
                String key = SlideReference.MakeKey(rtqp.DeckGuid, rtqp.SlideIndex);
                if (!slideImages.ContainsKey(key)) {
                    slideImages.Add(key, new SlideImage());
                }
                ((SlideImage)slideImages[key]).UpdateQuickPoll(rtqp);               
            }
            else {
                Debug.WriteLine("Slide Image Manager does not handle type: " + o.GetType().ToString());
            }
            //Notice that CP3 doesn't support scrolling mylar, so we don't bother with it here.
        }


        #endregion CP3

        #endregion Private Methods

        #region SlideReference Class

        /// <summary>
        /// Used to track current slide reference and other global properties
        /// </summary>
        private class SlideReference
        {
            //for CP
            private Guid deckGuid;
            private int slideIndex;
            //for RTDoc
            private Guid pageGuid;

            /// <summary>
            /// True if the reference has been set.
            /// </summary>
            public bool IsSet
            {
                get
                {
                    if ((pageGuid != Guid.Empty) || (deckGuid != Guid.Empty))
                        return true;
                    return false;
                }

            }

            //Other global settings
            private double size;
            private Color bgColor;

            public double Size
            {
                get { return size;  }
            }

            public Color BGColor
            {
                get { return bgColor; }
            }

            public SlideReference()
            {
                deckGuid = pageGuid = Guid.Empty;
                slideIndex = -1;
                size = 1.0;
                this.bgColor = Color.PapayaWhip;
            }

            public void SetCPReference(Guid deckGuid, int slideIndex)
            {
                pageGuid = Guid.Empty;
                this.deckGuid = deckGuid;
                this.slideIndex = slideIndex;
            }

            public void SetRTDocReference(Guid pageGuid)
            {
                deckGuid = Guid.Empty;
                slideIndex = -1;
                this.pageGuid = pageGuid;
            }

            public void SetSize(double size)
            {
                this.size = size;
            }

            /// <summary>
            /// Construct a hash key in the CP scenario.
            /// </summary>
            /// <param name="deckGuid"></param>
            /// <param name="slideIndex"></param>
            /// <returns></returns>
            public static string MakeKey(Guid deckGuid, int slideIndex)
            {
                return deckGuid.ToString() + "-" + slideIndex.ToString();
            }


            public override string ToString()
            {
                if (deckGuid != Guid.Empty)
                {
                    return "CP SlideReference: DeckGuid=" + deckGuid.ToString() + ";SlideIndex=" + slideIndex.ToString();
                }
                else
                {
                    return "RTDoc SlideReference: PageGuid=" + pageGuid.ToString();
                }
            }

            /// <summary>
            /// Return a string which we use as the hash code for the current reference.
            /// </summary>
            /// <returns></returns>
            public string GetStringHashCode()
            {
                if (deckGuid != Guid.Empty)
                {
                    return SlideReference.MakeKey(deckGuid, slideIndex);
                }
                else
                {
                    return pageGuid.ToString();
                }
            }

            public void SetBGColor(Color color)
            {
                this.bgColor = color;
            }
        }

        #endregion SlideReference Class

    }
}
