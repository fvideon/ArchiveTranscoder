using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using ArchiveTranscoder;
using System.Reflection;
using System.IO;
using XGetoptCS;


namespace ArchiveTranscoderCmd {
    class Program {
        private static ManualResetEvent m_JobCompleteEvent;
        private static bool m_MakeMobileVersion = false;
        private static bool m_Test = false;
        private static string m_Profile = null;
        private static string m_JobFile;
        private static string m_PreviousStatusMessage;

        static int Main(string[] args) {
            string errMsg = null;
            m_PreviousStatusMessage = "";

            Utility.DeleteExistingTempFilesAndDirs();

            errMsg = ParseArgs(args);
            if (errMsg != null) {
                ShowUsage(errMsg);
                return 1;
            }
            m_JobFile = args[args.Length - 1];

            ArchiveTranscoderBatch batch = ArchiveTranscoder.Utility.ReadBatchFile(m_JobFile);
            if (batch == null) {
                ShowUsage("Batch file is invalid: " + m_JobFile);
                return 2;
            }

            if (m_MakeMobileVersion) {
                errMsg = MakeMobileBatch(batch);
                if (errMsg != null) {
                    ShowUsage(errMsg);
                    return 7;
                }
            }

            ArchiveTranscoder.ArchiveTranscoder transcoder = new ArchiveTranscoder.ArchiveTranscoder();
            errMsg = transcoder.LoadJobs(batch);
            if (errMsg != null) {
                ShowUsage(errMsg);
                return 3;
            }

            //Respect database name and server as specified in the batch file
            if ((!String.IsNullOrEmpty(batch.DatabaseName)) && (!batch.DatabaseName.Equals(transcoder.DBName))) {
                transcoder.DBName = batch.DatabaseName;
            }
            if ((!String.IsNullOrEmpty(batch.Server)) && (!batch.Server.Equals(transcoder.SQLHost))) {
                transcoder.SQLHost = batch.Server;
            }

            if (!transcoder.VerifyDBConnection()) {
                ShowUsage("SQL Server connection failed.  Please set or verify SQL Server host and database name.");
                return 4;
            }

            //Warn about file overwriting
            if (transcoder.WillOverwriteFiles) {
                Console.WriteLine("Warning: Existing transcoder output will be overwritten.");
            }

            //Fail if decks don't exist or are unmatched.
            String deckWarnings = transcoder.GetDeckWarnings();
            if ((deckWarnings != null) && (deckWarnings != "")) {
                Console.WriteLine(deckWarnings);
                return 5;
            }

            if (m_Test) {
                transcoder.RunUnitTests();
                return 0;
            }

            transcoder.OnBatchCompleted += new ArchiveTranscoder.ArchiveTranscoder.batchCompletedHandler(transcoder_OnBatchCompleted);
            transcoder.OnStatusReport += new ArchiveTranscoder.ArchiveTranscoder.statusReportHandler(transcoder_OnStatusReport);
            m_JobCompleteEvent = new ManualResetEvent(false);

            errMsg = transcoder.Encode();
            if (errMsg != null) {
                Console.WriteLine(errMsg);
                return 6;
            }

            m_JobCompleteEvent.WaitOne();

            Console.WriteLine("Encoding completed.");
            Console.WriteLine("Output: " + GetJobOutputPath(batch));
            return 0; //success
        }

        /// <summary>
        /// Return the base name of the output for the first job in the batch.
        /// </summary>
        /// <param name="batch"></param>
        /// <returns></returns>
        private static string GetJobOutputPath(ArchiveTranscoderBatch batch) {
            string ret = "";
            if ((batch.Job.Length > 0) && 
                (batch.Job[0] != null)) {
                ret = Path.Combine(batch.Job[0].Path, batch.Job[0].BaseName);
            }
            return ret;
        }

        private static string MakeMobileBatch(ArchiveTranscoderBatch batch) {
            foreach (ArchiveTranscoderJob job in batch.Job) {
                job.BaseName = job.BaseName + "m";
                job.Target = new ArchiveTranscoderJobTarget[1];
                job.Target[0] = new ArchiveTranscoderJobTarget();
                job.Target[0].Type = "stream";
                job.Target[0].CreateAsx = "False";
                job.Target[0].CreateWbv = "False";
                if (m_Profile != null) {
                    FileInfo oldProfile = new FileInfo(job.WMProfile);
                    string newProfile = Path.Combine(oldProfile.DirectoryName, m_Profile + ".prx");
                    if (!File.Exists(newProfile)) {
                        return "Failed to find requested encoding profile: " + newProfile;
                    }
                    job.WMProfile = newProfile;
                }
                
     
                foreach (ArchiveTranscoderJobSegment segment in job.Segment) {
                    segment.VideoDescriptor = null;
                    segment.Flags = new string[1];
                    segment.Flags[0] = "SlidesReplaceVideo";
                }
            }

            return null;
        }

        /// <summary>
        /// We require the xml job file as the last argument.
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        private static string ParseArgs(string[] args) {
            if (args.Length == 0) {
                return "No job file specified.";
            }
            if (args.Length == 1) {
                return null;
            }

            char c;
            XGetopt go = new XGetopt();
            // Note: feed Getopt a string list of flag characters.  A character followed by a colon causes it to 
            // consume two args, placing the second in the Optarg property.
            // For example: "sa:v:" correctly parses "-s -a 3 -v 2"
            while ((c = go.Getopt(args.Length - 1, args, "mtp:")) != '\0') {
                //Console.WriteLine("Getopt returned '{0}'", c);
                switch (c) {
                    case 'm':
                        //Generate mobile version
                        m_MakeMobileVersion = true;
                        break;
                    case 't':
                        m_Test = true;
                        break;
                    case 'p':
                        m_Profile = go.Optarg;
                        break;
                    case '?':
                        return "Illegal or missing argument";
                }
            }

            //Enforce any constraints here.

            //Anything left over triggers an error too:
            if (go.Optarg != "") {
                return "Unexpected arguments: " + go.Optarg;
            }

            if ((m_Profile != null) && (!m_MakeMobileVersion)) {
                return "-p is only allowed with -m";
            }

            return null;
        }

        static void transcoder_OnStatusReport(string message) {
            if (m_PreviousStatusMessage.Equals(message))
                return;

            Console.WriteLine(message);
            m_PreviousStatusMessage = message;
        }

        static void transcoder_OnBatchCompleted() {
            m_JobCompleteEvent.Set();
        }

        private static void ShowUsage(String msg) {
            Assembly a = Assembly.GetExecutingAssembly();
            String exeName = Path.GetFileName(a.Location);
            string baseName = Path.GetFileNameWithoutExtension(a.Location);
            Console.WriteLine(msg);
            Console.WriteLine("Usage: " + exeName + " [-m [-p 'profile name']] [-t] job.xml");
            Console.WriteLine("  If -m, the tool will convert job parameters such that the video will be " +
                "created from the presentation stream (slides and ink).  The supplied xml file will not be modified."); 
            Console.WriteLine("  If -p, use the named encoding profile for the video created from presentation stream.");
            Console.WriteLine("  If -t, run unit tests.");

        }

    }
}
