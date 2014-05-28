using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using MSR.LST;
using System.IO;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.Collections;
using System.Drawing.Drawing2D;

namespace ArchiveTranscoder
{
    /// <summary>
    /// Given a slide or whiteboard reference, and optional slide image and sequence of ink operations, 
    /// return the raw bitmap of the slide including both background and ink as requested. Cache the image so that
    /// repeated requests can be satisfied with the minimum processing.  For now we have hardcoded to 320x240 output
    /// since target devices are portable devices with small screens.
    /// </summary>
    /// 
    /// Uniquely identifying ink strokes in Classroom Presenter:
    /// -Classroom Presenter maintains a table of stroke Guid to integer identifier, but the table is somewhat 
    /// redundant because the Guid is also stored in an extended property of the stroke indexed by a Guid constant.
    /// The table appears to exist for efficiency.
    /// -To add strokes, it looks up the int from the table and deletes any strokes it finds, then uses AddStrokesAtRectangle
    /// to add new ink, then looks in the new strokes extended property to find stroke guids and updates the guid to
    /// integer table for each.  Q: will the inbound ink ever contain more than one stroke?
    /// -To delete a stroke, it looks up the int identifier, then it uses Ink.CreateStrokes(int[]) to return a 
    /// strokes array, then Ink.DeleteStroke(stroke) to remove the zero'th element in the array.
    /// 
    /// RTDocuments: 
    /// -The guid to int lookup table is not used.  Instead the RTDocuments ink relies on the Guid in the extended property
    /// to identify strokes.  The property is indexed under the Guid static: RTStroke.ExtendedPropertyStrokeIdentifier.
    /// This is surely different from the one used by CP(??).
    /// 
    /// Since we are not a real-time app, we will use the RTDocument strategy of storing the Guid in the stroke's extended
    /// property.  We will arbitrarily choose to use the CP identifier.  Since we will be receiving both RTDocument and CP
    /// ink, we will explicitly set the extended property with each stroke received.
    /// Since it appears that inbound ink always contains one stroke, we will make that assumption explicit by throwing an 
    /// exception if it is not true.
    /// 
    /// Semi-transparent Ink:
    /// When we save ink to an image bitmap there appears to be no way to preserve its transparency.  This is an issue with
    /// highlighter pens.  To work around this, we save all the opaque ink in one collection and the non-opaque ink in a second
    /// collection.  We can generate the opaque ink bitmap once for all opaque strokes, but we need to generate a separate bitmap 
    /// overlay for each semi-transparent stroke in order to keep the particular stroke transparency value intact.
    /// 
    class SlideImage:IDisposable
    {
        #region Members

        Bitmap slideBitmap;
        byte[] rawCompositeImage;
        bool dirtyBit;
        Microsoft.Ink.Ink opaqueInk;
        Microsoft.Ink.Ink transparentInk;
        Dictionary<Guid, TextAnnotation> textAnnotations;
        Dictionary<Guid, DynamicImage> dynamicImages;
        string filePath;
        double scrollPos;
        double scrollExtent;
        double lastSize;
        Color lastBgColor;
        private QuickPoll quickPoll;

        /// <summary>
        /// Flag to indicate that the background image is a whiteboard and does not come from a PPT slide image.
        /// If this slide is a Student Submission or QuickPoll, the value of this flag should be inherited from 
        /// the association slide.
        /// </summary>
        private bool whiteboardBackground;

        #endregion Members

        #region Constructor

        public SlideImage()
        {
            whiteboardBackground = true;
            slideBitmap = null;
            rawCompositeImage = null;
            dirtyBit = true;
            opaqueInk = null;
            transparentInk = null;
            filePath = null;
            quickPoll = null;

            lastSize = 1.0;
            lastBgColor = Color.PapayaWhip;

            scrollPos = 0.0;
            scrollExtent = 1.5;
        }

        #endregion Constructor

        #region IDisposable Members

        public void Dispose()
        {
            if (slideBitmap != null)
            {
                //Note this must be explicitly disposed to prevent an exception when
                //SlideImageGenerator tries to delete the temp directory in its destructor.
                slideBitmap.Dispose();
                slideBitmap = null;
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Adding one stroke
        /// </summary>
        /// <param name="ink"></param>
        /// <param name="guid"></param>
        /// <param name="finished"></param>
        public void AddInk(Microsoft.Ink.Ink newInk, Guid guid)
        {
            if (newInk.Strokes.Count > 1) {
                AddInk(newInk);
            }
            else if (newInk.Strokes.Count > 0) {

                if (guid.Equals(Guid.Empty)) {
                    throw new ApplicationException("Ink Guid must not be Guid.Empty.");
                }

                if (newInk.Strokes[0].DrawingAttributes.Transparency != 255) {
                    //Ignore invisible strokes.  These can occur in CP when the user erases individual strokes.

                    if (newInk.Strokes[0].DrawingAttributes.Transparency == 0) {
                        AddInkToInk(ref opaqueInk, newInk, guid);
                    }
                    else {
                        AddInkToInk(ref transparentInk, newInk, guid);
                    }

                    dirtyBit = true;
                }
            }
        }



        /// <summary>
        /// Adding multiple strokes (which can't be erased individually because there are no Guids)
        /// Specifically, this is for student submission overlays.
        /// There are also some CP3 scenarios such as when the instructor opens a CP3 file that has pre-existing ink.
        /// </summary>
        /// <param name="ink"></param>
        public void AddInk(Microsoft.Ink.Ink newInk)
        {
            if ((newInk != null) && (newInk.Strokes.Count > 0))
            {
                //separate transparent and opaque strokes
                List<int> transparentStrokes = new List<int>();
                List<int> opaqueStrokes = new List<int>();

                //iterate over strokes, adding the ink ids to the correct list
                foreach (Microsoft.Ink.Stroke s in newInk.Strokes)
                {
                    if ((s.DrawingAttributes.Transparency != 0) && (s.DrawingAttributes.Transparency != 255))
                    {
                        transparentStrokes.Add(s.Id);
                    }
                    else
                    {
                        opaqueStrokes.Add(s.Id);
                    }
                }

                //Add transparent strokes to transparentInk.
                if (transparentStrokes.Count > 0)
                {
                    Microsoft.Ink.Strokes tStrokes = newInk.CreateStrokes((int[])transparentStrokes.ToArray());
                    if (transparentInk == null)
                    {
                        transparentInk = new Microsoft.Ink.Ink();
                    }
                    transparentInk.AddStrokesAtRectangle(tStrokes, tStrokes.GetBoundingBox());
                    dirtyBit = true;
                }

                //Add opaque strokes to opaqueInk.
                if (opaqueStrokes.Count > 0)
                { 
                    Microsoft.Ink.Strokes oStrokes = newInk.CreateStrokes((int[])opaqueStrokes.ToArray());
                    if (opaqueInk == null)
                    {
                        opaqueInk = new Microsoft.Ink.Ink();
                    }
                    opaqueInk.AddStrokesAtRectangle(oStrokes, oStrokes.GetBoundingBox());
                    dirtyBit = true;
                }
            }
        }

        /// <summary>
        /// Set the slide image.  We only support setting this once.
        /// We may modify the image and keep a reference to it, so the caller is responsible for cloning the image if needed.
        /// </summary>
        /// <param name="image"></param>
        public void SetImage(Image image, bool isWhiteboard)
        {
            if (image == null)
                return;

            //In some cases the image may be a Metafile.  We need a Bitmap.
            if (!(image is Bitmap))
            {
                slideBitmap = new Bitmap(image,960,720);
            }
            else
            {
                slideBitmap = (Bitmap)image;
            }
            whiteboardBackground = isWhiteboard;
            dirtyBit = true;
        }

        /// <summary>
        /// If a stroke exists with the guid in its extended properties, delete it.
        /// </summary>
        /// <param name="guid"></param>
        public void RemoveInk(Guid guid)
        {
            removeFromInk(opaqueInk,guid);
            removeFromInk(transparentInk, guid);
        }

        /// <summary>
        /// Removes ink AND text annotations from the slide
        /// </summary>
        public void RemoveAllInk()
        {
            if (opaqueInk != null)
            {
                opaqueInk.DeleteStrokes();
                dirtyBit = true;
            }
            if (transparentInk != null)
            {
                transparentInk.DeleteStrokes();
                dirtyBit = true;
            }
            if (textAnnotations != null) {
                textAnnotations.Clear();
                dirtyBit = true;
            }
        }

        public void SetScrollPos(double position, double extent)
        {
            if ((scrollPos != position) || (scrollExtent != extent))
            {
                scrollPos = position;
                scrollExtent = extent;
                dirtyBit = true;
            }
        }

        /// <summary>
        /// Set the path to an image file for this slide
        /// </summary>
        /// <param name="filePath"></param>
        public void SetFilePath(String filePath)
        {
            if (this.filePath != filePath)
            {
                this.filePath = filePath;
                dirtyBit = true;
            }
        }

        /// <summary>
        /// Return the raw bitmap representing the current state of this slide
        /// </summary>
        /// <param name="size"></param>
        /// <param name="bgColor"></param>
        /// <returns></returns>
        public BufferChunk GenerateFrame(double size, Color bgColor, int exportWidth, int exportHeight)
        {

            //If there have been no changes since we built the last bitmap, just return it:
            if ((this.rawCompositeImage != null) && (!dirtyBit) && (lastSize == size) && (lastBgColor == bgColor))
            {                
                return new BufferChunk(rawCompositeImage);
            }

            //build a new bitmap if needed.
            rawCompositeImage = null;
            if (slideBitmap == null)
            {
                //First try to build the image from a given file.  This will be a slide image.
                if ((filePath != null) && (File.Exists(filePath)))
                {
                    try
                    {
                        slideBitmap = (Bitmap)Image.FromFile(filePath);
                        whiteboardBackground = false;
                    }
                    catch
                    {
                        slideBitmap = null;
                    }
                }
            }

            //If the background did not derive from a slide image from PPT etc., then the whiteboardBackground flag should be
            //true.  This should work even if this slide is a student submission or quick poll because we will have inherited
            //the value from the association slide.
            if (whiteboardBackground)
            {
                //Generate a whiteboard background image if we don't already have one of the correct color.
                if ((slideBitmap == null) || (!bgColor.Equals(lastBgColor))) {
                    if (slideBitmap != null) {
                        slideBitmap.Dispose();
                    }
                    slideBitmap = (Bitmap)makeBlankImage(bgColor, 960, 720);
                }
            }


            rawCompositeImage = buildRawImage(slideBitmap, size, bgColor, exportWidth, exportHeight);

            if (rawCompositeImage != null)
            {
                lastSize = size;
                lastBgColor = bgColor;

                dirtyBit = false;
                return new BufferChunk(rawCompositeImage);
            }

            return null;
        }

        /// <summary>
        /// Return a copy of the slide image bitmap.  If none is defined, return a blank image with the given color.
        /// </summary>
        /// <returns></returns>
        internal Image GetSlideImage(Color bgColor, out bool isWhiteboard)
        {
            if (slideBitmap == null)
            {
                if ((filePath != null) && (File.Exists(filePath)))
                {
                    try
                    {
                        slideBitmap = (Bitmap)Image.FromFile(filePath);
                        whiteboardBackground = false;
                    }
                    catch
                    {
                        slideBitmap = null;
                    }
                }
            }

            if (slideBitmap == null)
                slideBitmap = (Bitmap)makeBlankImage(bgColor, 960, 720);

            isWhiteboard = whiteboardBackground;
            return (Image)slideBitmap.Clone();
        }

        internal void AddTextAnnotation(string text, Point origin, Guid guid, Font font, Color color, int height, int width) {
            if (textAnnotations == null) {
                textAnnotations = new Dictionary<Guid, TextAnnotation>();
            }
            if (textAnnotations.ContainsKey(guid)) {
                textAnnotations[guid] = new TextAnnotation(guid, text, color, font, origin, height, width);
            }
            else {
                textAnnotations.Add(guid, new TextAnnotation(guid, text, color, font, origin, height, width));
            }
            dirtyBit = true;
        }

        /// <summary>
        /// Obsolete
        /// </summary>
        /// <param name="guid"></param>
        internal void RemoveTextAnnotation(Guid guid) {
            if (textAnnotations == null) {
                return;
            }
            if (textAnnotations.ContainsKey(guid)) {
                textAnnotations.Remove(guid);
                dirtyBit = true;
            }
        }

        /// <summary>
        /// Remove annotations of type text or image.
        /// </summary>
        /// <param name="guid"></param>
        internal void RemoveAnnotation(Guid guid) {

            if ((textAnnotations != null) && (textAnnotations.ContainsKey(guid))) {
                textAnnotations.Remove(guid);
                dirtyBit = true;
            }

            if ((dynamicImages != null) && (dynamicImages.ContainsKey(guid))) {
                dynamicImages.Remove(guid);
                dirtyBit = true;
            }
            
        }

        internal void UpdateQuickPoll(ArchiveRTNav.RTQuickPoll rtqp) {
            if (quickPoll == null) {
                quickPoll = new QuickPoll();
            }
            quickPoll.Update((int)(rtqp.Style), rtqp.Results);
            dirtyBit = true;
        }


        internal void AddImageAnnotation(Guid id, int height, int width, Point origin, Image image) {
            if (dynamicImages == null) {
                dynamicImages = new Dictionary<Guid, DynamicImage>();
            }

            if (dynamicImages.ContainsKey(id)) {
                dynamicImages.Remove(id);
            }

            dynamicImages.Add(id, new DynamicImage(id, origin, width, height, image));
            dirtyBit = true;

        }


        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Add a stroke to an Ink
        /// </summary>
        /// <param name="destInk"></param>
        /// <param name="srcInk"></param>
        /// <param name="guid"></param>
        private void AddInkToInk(ref Microsoft.Ink.Ink destInk, Microsoft.Ink.Ink srcInk, Guid guid)
        {
            if (destInk == null)
            {
                //first stroke: create new ink
                destInk = new Microsoft.Ink.Ink();
            }

            //Remove the stroke, if it exists.
            RemoveInk(guid);

            //Set the extended property.
            if (srcInk.Strokes[0].ExtendedProperties.DoesPropertyExist(Constants.CPInkExtendedPropertyTag))
            {
                srcInk.Strokes[0].ExtendedProperties.Remove(Constants.CPInkExtendedPropertyTag);
            }
            srcInk.Strokes[0].ExtendedProperties.Add(Constants.CPInkExtendedPropertyTag, guid.ToString());

            //Add stroke
            destInk.AddStrokesAtRectangle(srcInk.Strokes, srcInk.Strokes.GetBoundingBox());
        }

        private void removeFromInk(Microsoft.Ink.Ink ink, Guid guid)
        {
            if (ink == null)
                return;

            foreach (Microsoft.Ink.Stroke s in ink.Strokes)
            {
                if (s.ExtendedProperties.DoesPropertyExist(Constants.CPInkExtendedPropertyTag))
                {
                    if (s.ExtendedProperties[Constants.CPInkExtendedPropertyTag].Data.Equals(guid.ToString()))
                    {
                        ink.DeleteStroke(s);
                        dirtyBit = true;
                        return;
                    }
                }
            }
        }


        /// <summary>
        /// Return a raw bitmap of the scaled image and optional ink overlay.
        /// We do not modify the image supplied.
        /// </summary>
        /// <param name="slideImage"></param>
        /// <param name="ink"></param>
        /// <returns></returns>
        private byte[] buildRawImage(Bitmap slideImage, double size, Color bgColor, int exportWidth, int exportHeight)
        {
            if (slideImage == null)
                throw new ApplicationException("Slide Image must not be null.");

            Image img = (Image)slideImage.Clone();

            //Scale the slide down if specified.
            if (size != 1.0)
            {
                scaleImage(ref img, size, bgColor);
            }

            //Add quickpoll overlay if applicable
            if (quickPoll != null) {
                addQuickPollOverlay(img, size);
            }

            addDynamicImages(img, size);

            AddOpaqueInkOverlay(size, img);

            addTransparentInkOverlay(img, size);

            addTextAnnotationOverlay(img, size);

            //Screen snapshots may have non-standard aspect ratio:
            fixAspectRatio(ref img);

            //WMFSDK wants the picture upside down:
            img.RotateFlip(RotateFlipType.RotateNoneFlipY);

            Bitmap bm; 

            if ((exportWidth == img.Width) && (exportHeight == img.Height))
            {
                //No scaling needed.  Ideally this will be the common case.
                bm = (Bitmap)img;
            }
            else if ((exportWidth == 0) || (exportHeight == 0) ||
                    (exportWidth == 320 && exportHeight == 240)) {
                //Export width or height not specified or 320x240.  Export default width and height.
                //GetThumbnailImage produces a nice scaled image up to about 320x240.
                bm = new Bitmap(img.GetThumbnailImage(320, 240, new Image.GetThumbnailImageAbort(abortCallback), IntPtr.Zero));
                img.Dispose();            
            }
            else
            {
                bm = scaleFullImage(img, exportWidth, exportHeight, Color.White);
            }

            //Generate the raw bitmap.
            return imageToRawBitmap(bm);
        }

        /// <summary>
        /// Scale the image to a new size and return it.  Dispose of the original.
        /// </summary>
        private Bitmap scaleFullImage(Image img, int newWidth, int newHeight, Color bgColor)
        {
            //Make a blank image of the new size
            Bitmap newImage = new Bitmap(newWidth, newHeight);

            Graphics g = Graphics.FromImage(newImage);
            GraphicsUnit gu = GraphicsUnit.Pixel;

            //scale the original image and draw it on the blank image
            RectangleF destRect = new RectangleF(0, 0, (float)newWidth, (float)newHeight);
            g.DrawImage(img, newImage.GetBounds(ref gu), img.GetBounds(ref gu), GraphicsUnit.Pixel);
            img.Dispose();
            g.Dispose();
            return newImage;
        }

        private void addDynamicImages(Image img, double size) {
            if ((dynamicImages != null) && (dynamicImages.Count > 0)) {
                Graphics g = Graphics.FromImage(img);
                foreach (DynamicImage di in dynamicImages.Values) {
                    float x = (float)di.Origin.X * (float)img.Width * (float)size / 500F;
                    float y = (float)di.Origin.Y * (float)img.Height * (float)size / 500F;
                    float h = (float)di.Height * (float)img.Height * (float)size / 500F;
                    float w = (float)di.Width * (float)img.Width * (float)size / 500F;
                    Rectangle bounds = new Rectangle((int)x, (int)y, (int)w, (int)h);
                    g.DrawImage(di.Img, bounds);
                }
                g.Dispose();
            }
        }


        private void AddOpaqueInkOverlay(double size, Image img) {
            //add the opaque ink overlay
            if ((opaqueInk != null) && (opaqueInk.Strokes.Count > 0)) {
                /// draw the slide data on a temporary graphics object in a temporary form
                System.Windows.Forms.Form tempForm = new System.Windows.Forms.Form();
                Graphics screenGraphics = tempForm.CreateGraphics();
                DibGraphicsBuffer dib = new DibGraphicsBuffer();
                Graphics tempGraphics = dib.RequestBuffer(screenGraphics, img.Width, img.Height);

                //Add the background color
                //First see if there is a Slide BG, if not, try the Deck.  Otherwise, use transparent.
                tempGraphics.DrawImage(img,0,0);

                //System.Drawing.Drawing2D.GraphicsState oldState = tempGraphics.Save();

                Microsoft.Ink.Renderer renderer = new Microsoft.Ink.Renderer();
                Matrix transformation = new Matrix();
                renderer.GetViewTransform(ref transformation);
                transformation.Scale(((float)img.Width/500f) * (float)size ,
                                     ((float)img.Height/500f) * (float)size);
                renderer.SetViewTransform(transformation);

                renderer.Draw(tempGraphics, opaqueInk.Strokes);
                
                //tempGraphics.Restore(oldState);

                Graphics toSave = Graphics.FromImage(img);
                dib.PaintBuffer(toSave, 0, 0);
                
                toSave.Dispose();
                tempGraphics.Dispose();
                dib.Dispose();
                screenGraphics.Dispose();
                tempForm.Dispose();
            }
        }


        private void AddOpaqueInkOverlayOld(double size, Image img) {
            //add the opaque ink overlay
            if ((opaqueInk != null) && (opaqueInk.Strokes.Count > 0)) {
                //Make a GIF Image from the ink.  Note that this image is assumed to be in a 500x500 pixel space.
                byte[] ba = opaqueInk.Save(Microsoft.Ink.PersistenceFormat.Gif);
                Image inkImg = Image.FromStream(new MemoryStream(ba));

                Graphics g = Graphics.FromImage(img);
                GraphicsUnit gu = GraphicsUnit.Pixel;

                //Get the origin from the ink Bounding Box (in ink space)
                Rectangle inkBB = opaqueInk.GetBoundingBox();

                //Convert the origin of the ink rectangle to pixel space (500x500)
                Microsoft.Ink.Renderer r = new Microsoft.Ink.Renderer();
                Point inkOrigin = inkBB.Location;
                r.InkSpaceToPixel(g, ref inkOrigin);

                //Adjust Y origin to account for horizontal scroll.  
                float scrolledYInkOrigin = (float)inkOrigin.Y - (500 * (float)scrollPos);

                //Scale and locate the destination rectangle of the ink within the slide image:
                RectangleF destRect = new RectangleF(
                    (float)inkOrigin.X * ((float)img.Width / 500f) * (float)size,
                    scrolledYInkOrigin * ((float)img.Height / 500f) * (float)size,
                    (float)inkImg.Width * ((float)img.Width / 500f) * (float)size,
                    (float)inkImg.Height * ((float)img.Height / 500f) * (float)size);

                //Draw the overlay:
                g.DrawImage(inkImg, destRect, inkImg.GetBounds(ref gu), GraphicsUnit.Pixel);

                g.Dispose();
            }
        }

        private void addQuickPollOverlay(Image img, double size) {
            Graphics g = Graphics.FromImage(img);

            Font writingFont = new Font(FontFamily.GenericSansSerif, 18.0f * (float)size);
            StringFormat format = new StringFormat(StringFormat.GenericDefault);
            format.Alignment = StringAlignment.Center;
            format.LineAlignment = StringAlignment.Center;

            int startX = (int)(img.Width * 0.5f * size);
            int endX = (int)(img.Width * 0.95f *size);
            int startY = (int)(img.Height * 0.3f * size);
            int endY = (int)(img.Height * 0.85f * size);
            int width = endX - startX;
            int height = endY - startY;
            RectangleF finalLocation = new RectangleF(startX, startY, width, height);

            List<string> names = quickPoll.GetNames();
            Dictionary<string, int> table = quickPoll.GetTable();

            // Draw the outline
            g.FillRectangle(Brushes.White, startX - 1, startY - 1, width, height);
            g.DrawRectangle(Pens.Black, startX - 1, startY - 1, width, height);

            // Count the total number of results
            int totalVotes = 0;
            foreach (int i in table.Values) {
                totalVotes += i;
            }

            // Draw the choices
            float columnWidth = width / names.Count;
            int columnStartY = (int)((height * 0.9f) + startY);
            int columnTotalHeight = columnStartY - startY;
            for (int i = 0; i < names.Count; i++) {
                // Draw the column
                int columnHeight = 0;
                if (totalVotes != 0) {
                    columnHeight = (int)Math.Round((float)columnTotalHeight * ((int)table[names[i]] / (float)totalVotes));
                }
                if (columnHeight == 0) {
                    columnHeight = 1;
                }
                g.FillRectangle(QuickPoll.ColumnBrushes[i], (int)(i * columnWidth) + startX, columnStartY - columnHeight, (int)columnWidth, columnHeight);

                // Draw the label
                g.DrawString(names[i],
                              writingFont,
                              Brushes.Black,
                              new RectangleF((i * columnWidth) + startX, columnStartY, columnWidth, endY - columnStartY),
                              format);

                // Draw the number
                string percentage = String.Format("{0:0%}", (totalVotes == 0) ? 0 : (float)(((int)table[names[i]] / (float)totalVotes)));
                int numberHeight = (endY - columnStartY) * 2;
                RectangleF numberRectangle = new RectangleF((i * columnWidth) + startX,
                                                             (numberHeight > columnHeight) ? (columnStartY - columnHeight - numberHeight) : (columnStartY - columnHeight),
                                                             columnWidth,
                                                             numberHeight);
                string numberString = percentage + System.Environment.NewLine + "(" + table[names[i]].ToString() + ")";
                g.DrawString(numberString, writingFont, Brushes.Black, numberRectangle, format);
            }
        
        }

        private void addTextAnnotationOverlay(Image img, double size) {
            if ((textAnnotations != null) && (textAnnotations.Count > 0)) {
                Graphics g = Graphics.FromImage(img);
                foreach (TextAnnotation ta in textAnnotations.Values) {
                    float x = (float)ta.Origin.X * (float)img.Width * (float)size / 500F;
                    float y = (float)ta.Origin.Y * (float)img.Height * (float)size / 500F;
                    float fontSize = ta.Font.Size * (float)img.Width * (float)size / 500F;
                    Font f = new Font(ta.Font.FontFamily, fontSize, ta.Font.Style);
                    //I'm not totally sure that CP3.0 archives include valid width and height.  
                    //They started being used to define a bounding box in 3.1.
                    if ((ta.Height > 0) && (ta.Width > 0)) {
                        //If width and height seem valid, use them:
                        float h = (float)ta.Height * (float)img.Height * (float)size / 500F;
                        float w = (float)ta.Width * (float)img.Width * (float)size / 500F;
                        Rectangle bounds = new Rectangle((int)x, (int)y, (int)w, (int)h);
                        g.DrawString(ta.Text, f, new SolidBrush(ta.Color), bounds);
                    }
                    else { 
                        //If no width or height, just draw to the right from the origin.
                        g.DrawString(ta.Text, f, new SolidBrush(ta.Color),x,y);                 
                    }
                }
                g.Dispose();
            }
        }

        private void addTransparentInkOverlay(Image img, double size)
        {
            //Ink transparency is in the range 0-255 with 0==opaque and 255==invisible.
            //In practice CP sets transparency to 160, but we should be prepared for
            //arbitrary values.
            //We have no assurance that individual strokes won't have different transparencies, so we'll
            //do each stroke as a separate overlay.

            if ((transparentInk != null) && (transparentInk.Strokes.Count > 0))
            {
                foreach (Microsoft.Ink.Stroke s in transparentInk.Strokes)
                {
                    Microsoft.Ink.Strokes strokes = transparentInk.CreateStrokes(new int[] {s.Id});
                    addTransparentStroke(img, strokes, size);
                }
            }
        }

        /// <summary>
        /// Overlay one stroke on the image.
        /// </summary>
        /// <param name="img"></param>
        /// <param name="s"></param>
        /// <param name="size"></param>
        /// PRI2: I believe we can use the DibGraphicsBuffer (see opaque ink handling above) to improve the 
        /// look of the transparency.  This currently has the following shortcoming:  The colors produced by the exported gif
        /// from CP's highlight pens have some whiteness which causes a little fog over the dark colors
        /// below, for example the yellow highlight over a black background results in a lighter dark color, 
        /// no longer quite black.  In contrast the native ink mixing gives us nearly the original black.  It is as if
        /// the colors are on colored but clear mylar that overlay one another.  It's probably possible to get this effect
        /// with DrawImage somehow??
        private void addTransparentStroke(Image img, Microsoft.Ink.Strokes s, double size)
        {
            Microsoft.Ink.Ink tmpInk = new Microsoft.Ink.Ink();
            tmpInk.AddStrokesAtRectangle(s,s.GetBoundingBox());
            //Make a GIF Image from the Stroke.  Note that this image is assumed to be in a 500x500 pixel space.
            byte[] ba = tmpInk.Save(Microsoft.Ink.PersistenceFormat.Gif);
            Image inkImg = Image.FromStream(new MemoryStream(ba));

            Graphics g = Graphics.FromImage(img);

            //Get the origin from the ink Bounding Box (in ink space)
            Rectangle inkBB = tmpInk.GetBoundingBox();

            //Convert the origin of the ink rectangle to pixel space (500x500)
            Microsoft.Ink.Renderer r = new Microsoft.Ink.Renderer();
            Point inkOrigin = inkBB.Location;
            r.InkSpaceToPixel(g, ref inkOrigin);

            //Convert the transparency coefficient from 0-255 with 0==opaque to the range of 0-1 with 1==opaque.
            int t1 = Math.Abs(tmpInk.Strokes[0].DrawingAttributes.Transparency - 255);
            float t2 = (float)t1 / 255f;
            
            //Setting transparency with a ColorMatrix
            float[][] ptsArray ={ 
                new float[] {1, 0, 0, 0, 0}, //r
                new float[] {0, 1, 0, 0, 0}, //g
                new float[] {0, 0, 1, 0, 0}, //b
                new float[] {0, 0, 0, t2, 0}, //alpha
                new float[] {0, 0, 0, 0, 1}};
            ColorMatrix clrMatrix = new ColorMatrix(ptsArray);
            ImageAttributes imgAttributes = new ImageAttributes();
            imgAttributes.SetColorMatrix(clrMatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);

            //Adjust Y origin to account for horizontal scroll.  (scrollPos becomes more positive as ink goes upward.)
            float scrolledYInkOrigin = (float)inkOrigin.Y - (500 * (float)scrollPos); //Still in 500x500 space

            //Scale and locate the destination rectangle of the ink within the slide image:
            RectangleF destRect = new RectangleF(
                (float)inkOrigin.X * ((float)img.Width / 500f) * (float)size,
                scrolledYInkOrigin * ((float)img.Height / 500f) * (float)size,
                (float)inkImg.Width * ((float)img.Width / 500f) * (float)size,
                (float)inkImg.Height * ((float)img.Height / 500f) * (float)size);

            Rectangle destRect2 = new Rectangle((int)destRect.X, (int)destRect.Y, (int)destRect.Width, (int)destRect.Height);

            //Draw the overlay:
            g.DrawImage(inkImg, destRect2, 0,0,inkImg.Width,inkImg.Height, GraphicsUnit.Pixel, imgAttributes);
            g.Dispose();
        }

        /// <summary>
        /// Scale the image down and place it in the upper left of a blank image.
        /// </summary>
        /// <param name="img"></param>
        /// <param name="size"></param>
        /// <param name="bgColor"></param>
        private void scaleImage(ref Image img, double size, Color bgColor)
        {
            //Make a blank image of the same size as the original
            Image blankImage = makeBlankImage(bgColor, img.Width, img.Height);

            Graphics g = Graphics.FromImage(blankImage);
            GraphicsUnit gu = GraphicsUnit.Pixel;

            //scale the original image and draw it in the upper left of the blank image
            RectangleF destRect = new RectangleF(0, 0, (float)img.Width * (float)size, (float)img.Height * (float)size);
            g.DrawImage(img, destRect, blankImage.GetBounds(ref gu), GraphicsUnit.Pixel);
            img.Dispose();
            img = blankImage;
            g.Dispose();
        }

        /// <summary>
        /// Return a bitmap of the background color
        /// </summary>
        /// <param name="bgColor"></param>
        /// <returns></returns>
        private Image makeBlankImage(Color bgColor, int w, int h)
        {
            Image img = new Bitmap(w, h);
            SolidBrush bgBrush = new SolidBrush(bgColor);
            Graphics g = Graphics.FromImage(img);
            g.FillRectangle(bgBrush, new Rectangle(0, 0, w, h));
            g.Dispose();
            bgBrush.Dispose();
            return img;
        }

        /// <summary>
        /// Return the raw uncompressed RGB24 representation of the image.
        /// Do not modify the input Bitmap.
        /// </summary>
        /// <param name="img"></param>
        /// <returns></returns>
        private byte[] imageToRawBitmap(Bitmap img)
        {
            //convert pixel format to RGB24
            //Should we check to make sure it isn't already?
            Bitmap rgbbm = img.Clone(new Rectangle(0, 0, img.Width, img.Height), PixelFormat.Format24bppRgb);
            int bytesPerPixel = 3;

            //Extract a byte[] from the Bitmap
            BitmapData bd = rgbbm.LockBits(new Rectangle(0, 0, img.Width, img.Height), ImageLockMode.ReadOnly, rgbbm.PixelFormat);
            int cbuf = img.Width * img.Height * bytesPerPixel;
            byte[] rawdata = new byte[cbuf];
            Marshal.Copy(bd.Scan0, rawdata, 0, cbuf);
            rgbbm.UnlockBits(bd);

            //Explicit disposal used to be necessary, but the GC seems to be working better now.  What the heck, do it anyway.
            rgbbm.Dispose();

            return rawdata;
        }

        private bool abortCallback()
        {
            return false;
        }

        /// <summary>
        /// Screen shots produce images with non-standard aspect ratio.  If this image has a non-standard 
        /// aspect ratio, draw it into an image that has a standard 1.33 ratio, return the result, and 
        /// return true.  Otherwise return false.
        /// </summary>
        /// <param name="img"></param>
        /// <param name="jpgFile"></param>
        /// <returns></returns>
        private bool fixAspectRatio(ref Image img)
        {
            double ar = ((double)img.Width) / ((double)img.Height);
            if (ar < 1.32)
            {
                //in this case the image is not wide enough.
                double w = ((double)img.Height) * 1.333;
                int newWidth = (int)Math.Round(w);
                Bitmap b = new Bitmap(newWidth, img.Height);
                Graphics g = Graphics.FromImage(b);
                SolidBrush whiteBrush = new SolidBrush(Color.White);
                //make a white background
                g.FillRectangle(whiteBrush, 0, 0, b.Width, b.Height);
                //paste the original image on the top of the new bitmap.
                g.DrawImage(img, 0, 0, img.Width, img.Height);

                whiteBrush.Dispose();
                img.Dispose();
                g.Dispose();
                img = b;
                return true;
            }
            else if (ar > 1.34)
            {
                //in this case the image is not tall enough.
                double h = ((double)img.Width) / 1.333;
                int newHeight = (int)Math.Round(h);
                Bitmap b = new Bitmap(img.Width, newHeight);
                Graphics g = Graphics.FromImage(b);
                SolidBrush whiteBrush = new SolidBrush(Color.White);
                //make a white background
                g.FillRectangle(whiteBrush, 0, 0, b.Width, b.Height);
                //paste the original image on the top of the new bitmap.
                g.DrawImage(img, 0, 0, img.Width, img.Height);

                whiteBrush.Dispose();
                img.Dispose();
                g.Dispose();
                img = b;
                return true;
            }
            return false;
        }

        //PRI2: we don't currently distinguish between missing slides and whiteboards.  The slide type information
        // is available at least in Classroom Presenter scenarios.

        ///// <summary>
        ///// Return an image with the words "Slide Missing" across the middle.
        ///// Create the image once the first time it is used.
        ///// </summary>
        ///// <returns></returns>
        //private Image getMissingSlideImage()
        //{
        //    if (missingSlideImage == null)
        //    {
        //        Image img = new Bitmap(320, 240);
        //        SolidBrush bgBrush = new SolidBrush(Color.Wheat);
        //        Graphics g = Graphics.FromImage(img);
        //        g.FillRectangle(bgBrush, new Rectangle(0, 0, 320, 240));
        //        GraphicsUnit gu = GraphicsUnit.Pixel;
        //        RectangleF bounds = img.GetBounds(ref gu);
        //        SolidBrush myBrush = new SolidBrush(Color.Gray);
        //        Font myFont = new Font("Ariel", 30);
        //        g.DrawString("Slide Missing", myFont, myBrush, 38, 100);
        //        g.Dispose();
        //        myBrush.Dispose();
        //        myFont.Dispose();
        //        missingSlideImage = img;
        //    }
        //    return (Image)missingSlideImage.Clone();
        //}

        #endregion Private Methods

    }

    #region TextAnnotation Class

    public class TextAnnotation {
        private Point origin;
        public Point Origin {
            get { return origin; }
            set { origin = value; }
        }

        private Color color;
        public Color Color {
            get { return color; }
            set { color = value; }
        }

        private Font font;
        public Font Font {
            get { return font; }
            set { font = value; }
        }

        private String text;
        public String Text {
            get { return text; }
            set { text = value; }
        }

        private Guid id;
        public Guid Id {
            get { return id; }
            set { id = value; }
        }

        private int height;
        private int width;

        public int Height {
            get { return height; }
            set { height = value; }
        }

        public int Width {
            get { return width; }
            set { width = value; }
        }

        public TextAnnotation(Guid id, String text, Color color, Font font, Point origin, int height, int width) {
            this.id = id;
            this.text = text;
            this.color = color;
            this.font = font;
            this.origin = origin;
            this.height = height;
            this.width = width;
        }
    }

    #endregion TextAnnotation Class

    #region DynamicImage Class

    public class DynamicImage {
        private Point origin;
        public Point Origin {
            get { return origin; }
            set { origin = value; }
        }

        private Guid id;
        public Guid Id {
            get { return id; }
            set { id = value; }
        }

        private int width;
        public int Width {
            get { return width; }
            set { width = value; }
        }

        private int height;
        public int Height {
            get { return height; }
            set { height = value; }
        }

        private Image image;
        public Image Img {
            get { return this.image; }
            set { this.image = value; }
        }

        public DynamicImage(Guid id, Point origin, int width, int height, Image image) {
            this.id = id;
            this.origin = origin;
            this.width = width;
            this.height = height;
            this.image = image;
        }
    }

    #endregion DynamicImage Class

    #region QuickPoll Class

    public class QuickPoll {
        private bool m_Enabled;
        private QuickPollStyle m_QuickPollStyle;
        private int[] m_Results;

        public QuickPoll() {
            m_Enabled = false;
            m_Results = new int[0];
            m_QuickPollStyle = QuickPollStyle.ABCD;
        }

        /// <summary>
        /// Make the QuickPoll visible or invisible
        /// </summary>
        public bool Enabled {
            get { return m_Enabled; }
            set { m_Enabled = value; }
        }

        public List<string> GetNames() {
            return QuickPoll.GetVoteStringsFromStyle(this.m_QuickPollStyle);
        }

        /// <summary>
        /// Return the current QuickPoll table
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, int> GetTable() {
            Dictionary<string, int> ret = new Dictionary<string, int>();
            lock (this) {
                List<string> names = GetNames();
                int index = 0;
                foreach (string s in names) {
                    if (m_Results.Length > index) {
                        ret.Add(s, m_Results[index]);
                    }
                    else {
                        ret.Add(s, 0);
                    }
                    index++;
                }
            }
            return ret;
        }

        /// <summary>
        /// Apply new results to the QuickPoll
        /// </summary>
        /// <param name="results"></param>
        public void Update(int styleAsInt, int[] results) {
            lock (this) {
                m_QuickPollStyle = (QuickPollStyle)styleAsInt;
                m_Results = results;
            }
        }


        #region Statics

        /// <summary>
        /// A helper member that specifies the brushes to use for the various columns
        /// </summary>
        public static System.Drawing.Brush[] ColumnBrushes = new System.Drawing.Brush[] { Brushes.Orange, 
                                                                            Brushes.Cyan, 
                                                                            Brushes.Magenta, 
                                                                            Brushes.Yellow, 
                                                                            Brushes.GreenYellow };

        /// <summary>
        /// Get the heading strings for a QuickPoll Style.
        /// </summary>
        /// <param name="style"></param>
        /// <returns></returns>
        public static List<string> GetVoteStringsFromStyle(QuickPollStyle style) {
            List<string> strings = new List<string>();
            switch (style) {
                case QuickPollStyle.YesNo:
                    strings.Add("Yes");
                    strings.Add("No");
                    break;
                case QuickPollStyle.YesNoBoth:
                    strings.Add("Yes");
                    strings.Add("No");
                    strings.Add("Both");
                    break;
                case QuickPollStyle.YesNoNeither:
                    strings.Add("Yes");
                    strings.Add("No");
                    strings.Add("Neither");
                    break;
                case QuickPollStyle.ABC:
                    strings.Add("A");
                    strings.Add("B");
                    strings.Add("C");
                    break;
                case QuickPollStyle.ABCD:
                    strings.Add("A");
                    strings.Add("B");
                    strings.Add("C");
                    strings.Add("D");
                    break;
                case QuickPollStyle.ABCDE:
                    strings.Add("A");
                    strings.Add("B");
                    strings.Add("C");
                    strings.Add("D");
                    strings.Add("E");
                    break;
                case QuickPollStyle.ABCDEF:
                    strings.Add("A");
                    strings.Add("B");
                    strings.Add("C");
                    strings.Add("D");
                    strings.Add("E");
                    strings.Add("F");
                    break;
                case QuickPollStyle.Custom:
                    // Do Nothing for now
                    break;
            }

            return strings;
        }

        #endregion Statics

        #region Enum

        public enum QuickPollStyle {
            Custom = 0,
            YesNo = 1,
            YesNoBoth = 2,
            YesNoNeither = 3,
            ABC = 4,
            ABCD = 5,
            ABCDE = 6,
            ABCDEF = 7
        }

        #endregion Enum
    }
    #endregion QuickPoll Class

}
