using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;

namespace ImageFilter {
    public class ImageFilter {

        public static void Validate() {
            if (String.IsNullOrEmpty(Config.ImageMagickPath)) {
                throw new ApplicationException("ImageMagick path must be specified in app.config.");
            }
        }

        private bool done = false;
        private string imageMagickPath = null;
        private double differenceThreshold = 500.0;
        private string fileGlob = "*.*";
        private bool debug = false;
        private IComparer<string> filenameSortComparer = new DefaultComparer();
        
        /// <summary>
        /// High enough value to indicate that the images differ.  Return this
        /// in case we fail to get the real metric.
        /// </summary>
        private const double DEFAULT_DIFFERENCE_METRIC = 10000.0;

        /// <summary>
        /// Causes the filter method to emit verbose output.
        /// </summary>
        public bool Debug {
            get { return this.debug;  }
            set { this.debug = value; }
        }

        /// <summary>
        /// ImageMagick compare (MAE metric) values greater or equal to this
        /// are taken to mean that the images differ enough.  The default is 500.0.
        /// </summary>
        public double DifferenceThreshold {
            get { return this.differenceThreshold; }
            set { this.differenceThreshold = value; }
        }

        /// <summary>
        /// The wildcard string that identifies the source image files in the input directory.
        /// Examples: "*.jpg", "capture*.png", "*.*".  The default is "*.*".
        /// </summary>
        public string FileGlob {
            get { return this.fileGlob; }
            set { this.fileGlob = value; }
        }

        /// <summary>
        /// Comparer to use for sorting filenames.  This is used prior to the filter operation.
        /// The default comparer is a simple lexicographical sort.
        /// </summary>
        public IComparer<string> FilenameSortComparer {
            get { return this.filenameSortComparer; }
            set { this.filenameSortComparer = value; }
        }

        public class DefaultComparer : IComparer<string> {
            public int Compare(string x, string y) {
                return x.CompareTo(y);
            }
        }

        /// <summary>
        /// specific to sorting files like this "slide1.jpg", "slide2.jpg", "slide10.jpg", etc
        /// </summary>
        public class PMPComparer : IComparer<string> {
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
        /// Alternate constructor using app.config to determine paths
        /// </summary>
        public ImageFilter() : this(Config.ImageMagickPath, Config.TempPath) { }

        /// <summary>
        /// Alternate constructor using app.config to determine only the ImageMagick path
        /// </summary>
        /// <param name="tempPath"></param>
        public ImageFilter(string tempPath) : this(Config.ImageMagickPath, tempPath) { }

        /// <summary>
        /// Verify that required external Imagemagick tools are available. 
        /// The path should be to the bin directory which contains at least the 
        /// compare.exe tool
        /// Throw an exception if there are any problems found.  
        /// </summary>
        /// <param name="imageMagickPath">Path to ImageMagick bin directory.</param>
        /// <param name="tempPath">Override to default or app.config temp path parameter
        /// Set to null to use default or app.config value. </param>
        public ImageFilter(string imageMagickPath, string tempPath) { 
            //Validate imagemagick path
            if (!File.Exists(Path.Combine(imageMagickPath,"compare.exe"))) {
                throw new ArgumentException("Failed to find compare.exe at the ImageMagick path: " + imageMagickPath);
            }

            if (!String.IsNullOrEmpty(tempPath)) {
                if (Directory.Exists(tempPath)) {
                    Config.TempPath = tempPath;
                }    
            }

            this.imageMagickPath = imageMagickPath;
        }

        /// <summary>
        /// copy from inputDir to outputDir those images that are different enough from
        /// their predecessor.  
        /// </summary>
        /// <param name="inputDir">Source directory of images</param>
        /// <param name="outputDir">Newly created directory of filtered images</param>
        /// <param name="result">Details/errors about the filter process -- much more verbose if Debug is true</param>
        /// <returns>False for any error</returns>
        public bool Filter(string inputDir, out string outputDir, out string result) {
            this.done = false;
            outputDir = "";
            result = "";
            if (String.IsNullOrEmpty(this.imageMagickPath)) {
                result = "ImageMagick path is not valid.";
                return false;
            }

            //Make sure the input dir exists and contains at least one image
            if (!checkInputDir(inputDir, out result)) {
                return false;
            }

            //Create an output dir
            if (!createOutputDir(out outputDir, out result)) {
                return false;
            }

            //read/compare/copy from input to output
            return filter(inputDir, outputDir, out result);

        }

        /// <summary>
        /// Return true if the images are more different than the threshold permits,
        /// otherwise false.  Return true if one of the files is null.
        /// Return false if both files are null.
        /// Also return false for any error.
        /// Result hold details about the comparison.
        /// </summary>
        /// <param name="f1"></param>
        /// <param name="f2"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        public bool ImagesDiffer(string f1, string f2, out string result, out bool error, out double metric) {
            result = null;
            error = false;
            metric = DEFAULT_DIFFERENCE_METRIC;
            try {
                return imagesDiffer(f1, f2, out result, out error, out metric);
            }
            catch (Exception e) {
                result = e.ToString();
                error = true;
                return false;
            }
        }

        private bool filter(string inputDir, string outputDir, out string result) {
            result = "";
            string debugmsg = "";
            bool retSuccess = true;
            try {
                string[] files = Directory.GetFiles(inputDir, this.fileGlob);
                Array.Sort(files, this.filenameSortComparer);
                string previous = null;
                foreach (string file in files) {
                    string msg;
                    if (previous != null) {
                        debugmsg += "compare file1: " + previous + ";";
                    }
                    if (file != null) {
                        debugmsg += "compare file2: " + file + ";";
                    }
                    bool error = false;
                    double metric;
                    if (imagesDiffer(previous, file, out msg, out error, out metric)) {
                        File.Copy(file, Path.Combine(outputDir, Path.GetFileName(file)));
                        debugmsg += " Image files differ;";
                        previous = file;
                    }
                    else {
                        debugmsg += " Image files are the same;";
                    }
                    if (error) retSuccess = false;
                    debugmsg += msg + "\r\n";
                }
                if (debug) result = debugmsg;
                return retSuccess;
            }
            catch (Exception e) {
                result = e.ToString();
                return false;
            }
        }

        /// <summary>
        /// Compare two files.  If one is null but not the other, return true.
        /// If both are null return false.  Otherwise use the external tool to 
        /// determine if they are dissimiliar enough.
        /// Can throw an exception in the case of an error reading files.
        /// </summary>
        /// <param name="previous"></param>
        /// <param name="file"></param>
        /// <returns>true if they differ, false if they are alike</returns>
        private bool imagesDiffer(string previous, string file, out string dbgmsg, out bool error, out double metric) {
            error = false;
            metric = DEFAULT_DIFFERENCE_METRIC;
            dbgmsg = "";
            if (((previous == null) && (file != null)) ||
                ((previous!=null) && (file == null))) {
                dbgmsg = "Returning true because one input is null.";
                return true;
            }

            if ((previous == null) && (file == null)) {
                dbgmsg = "Returning false because both inputs are null.";
                return false;
            }
 
            string msg;
            metric = imageMagickCompare(Path.GetFullPath(previous), Path.GetFullPath(file), out msg, out error);
            dbgmsg = "compare returned " + metric.ToString() + "; ";
            dbgmsg += "compare message: " + msg + "; ";
            return (metric >= this.differenceThreshold);

        }

        /// <summary>
        /// Return a double value representing the difference between the two files.
        /// Can throw and exception in case of an error reading files.
        /// If ImageMagick returns an error, assume the files differ.
        /// </summary>
        /// <param name="previous"></param>
        /// <param name="file"></param>
        /// <returns></returns>
        private double imageMagickCompare(string previous, string file, out string debugmsg, out bool error) {
            error = false;
            //compare –metric MAE img1.jpg img2.jpg null:
            string cmd = Path.Combine(this.imageMagickPath,"compare.exe");
            string args = " -metric MAE \"" + previous + "\" \"" + file + "\" null.jpg";
            string result;
            debugmsg = "";
            if (runProcess(cmd, args, out result)) {
                debugmsg = "compare returned result: " + result + "; stdErr: " + stdErr;
                double d = parseMagickOutput(stdErr, out error);
                if (error) debugmsg += "; Failed to parse Imagemagick output.";
                return d;
            }
            else {
                debugmsg = "compare process failed: " + result + ";stdErr:" + stdErr + ";stdOut:" + stdOut;
                error = true;
                return DEFAULT_DIFFERENCE_METRIC;
            }
        }

        private double parseMagickOutput(string stdOut, out bool error) {
            //Get the first word of the output
            error = false;
            try {
                string[] words = stdOut.Split(new char[] { ' ' });
                if (words.Length > 0) {
                    double d;
                    if (double.TryParse(words[0], out d)) {
                        return d;
                    }
                }
                error = true;
            }
            catch {
                error = true;
            }
            return DEFAULT_DIFFERENCE_METRIC;
        }


        private bool createOutputDir(out string outputDir, out string result) {
            outputDir = "";
            result = "";
            try {
                outputDir = Path.Combine(Config.TempPath, Guid.NewGuid().ToString());
                Directory.CreateDirectory(outputDir);
                return true;
            }
            catch (Exception e) {
                result = "Failed to create output directory: " + e.ToString();
                outputDir = "";
                return false;
            }
        }

        /// <summary>
        /// Make sure the directory exists and contains at least one .jpg file.
        /// </summary>
        /// <param name="inputDir"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        private bool checkInputDir(string inputDir, out string result) {
            result = "";
            try {
                if (Directory.Exists(inputDir)) {
                    string[] files = Directory.GetFiles(inputDir, this.fileGlob);
                    if ((files == null) || (files.Length == 0)) {
                        result = "Input directory: " + inputDir + " contains no files matching the pattern: " + this.fileGlob;
                        return false;
                    }
                }
                else {
                    result = "Input directory does not exist: " + inputDir;
                    return false;
                }
            }
            catch (Exception e) {
                result = "Exception while checking input directory: " + inputDir + "; " + e.ToString();
                return false;
            }
            return true;
        }

        #region Process Control

        private string stdErr;
        private string stdOut;

        private bool runProcess(string command, string arguments, out string result) {
            //Setup and start the process
            result = "";
            stdErr = "";
            stdOut = "";

            Process p = new Process();
            try {
                p.StartInfo.FileName = command;
                p.StartInfo.Arguments = arguments;
                p.StartInfo.WorkingDirectory = System.IO.Path.GetTempPath();
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardError = true;
                p.StartInfo.RedirectStandardOutput = true;
                p.OutputDataReceived += new DataReceivedEventHandler(outputDataReceived);
                p.ErrorDataReceived += new DataReceivedEventHandler(ErrorDataReceived);
                p.StartInfo.CreateNoWindow = true;
                p.Start();
                p.BeginOutputReadLine();
                p.BeginErrorReadLine();
            }
            catch (Exception e) {
                //Process failed to start.
                result = e.ToString();
                return false;
            }

            while (!this.done) {
                if (p.WaitForExit(1000)) {
                    //WaitForExit returns true iif timeout did not elapse.
                    if (p.ExitCode == 0) { //success
                        // There is a bug somewhere around here that sometimes causes us to miss the last line of output.
                        // This is also noted in on-line forums.  It seems that the WaitforExit can return before the last
                        // OutputDataReceived.  One suggested solution is to try the process "Exited" event.
                        // Another is to wait for Output Data Received to be raised with data null or empty, and
                        // use that to determine when the process is done.  More testing is needed.
                        return true;
                    }
                    else if (p.ExitCode == 1) {
                        // For some versions of imagemagick compare this exit code seems to indicate just that the images are
                        // different.
                        result = "Process returned exit code: 1";
                        return true;
                    }

                    //If a batch script returns with "exit N" (where N is an int), the N lands in ExitCode.
                    result = "Process failed with exit code:" + p.ExitCode.ToString();
                    return false;
                }
                else {
                    //Check for thread terminate, then continue
                    if (this.done) {
                        p.Kill();
                        result = "Process terminated by owner.";
                        return false;
                    }
                }
            }
            return false;
        }

        private void ErrorDataReceived(object sender, DataReceivedEventArgs e) {
            if (e.Data != null) {
                //Debug.WriteLine("Process error: " + e.Data);
                stdErr += e.Data + "\r\n";
            }
        }

        private void outputDataReceived(object sender, DataReceivedEventArgs e) {
            if (e.Data != null) {
                //Debug.WriteLine("Process output: " + e.Data);
                stdOut += e.Data + "\r\n";
            }
        }

        #endregion Process Control

    }

    
}
