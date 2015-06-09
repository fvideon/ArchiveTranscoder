using System;
using MSR.LST.Net.Rtp;
using MSR.LST;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;
using System.Drawing;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Test code
	/// </summary>
	public class Test
	{
		public Test() { }

        #region Test Code

        public static void CheckAzureDBConnectivity() {
            DatabaseUtility.SQLConnectionString = "Data Source=tczxnf5cbu.database.windows.net;Initial Catalog=ArchiveService;Integrated Security=false;User ID=cxp;Password=ConferenceXP!";
            if (DatabaseUtility.CheckConnectivity()) {
                Debug.WriteLine("Azure database is alive.");
                Conference[] confs = DatabaseUtility.GetConferences();
            }
            else {
                Debug.WriteLine("Azure database is not alive.");
            }
        }

        /// <summary>
        /// Read from a stream and attempt to find various issues.
        /// </summary>
        /// <param name="streamId"></param>
        /// <param name="durationSeconds"></param>
        public static void VerifyStream(int streamId, int durationSeconds) {
            DBStreamPlayer streamPlayer = new DBStreamPlayer(streamId, (long)durationSeconds * (long)Constants.TicksPerSec, PayloadType.dynamicVideo);
            Console.WriteLine("VerifyStream: streamID=" + streamId.ToString() + 
                "; requested duration Ticks=" + ((long)durationSeconds * (long)Constants.TicksPerSec).ToString());
            verifyStream(streamPlayer);
        }

         public static void VerifyStream(int streamId, DateTime startTime, DateTime endTime) {
            DBStreamPlayer streamPlayer = new DBStreamPlayer(streamId, startTime.Ticks, endTime.Ticks, PayloadType.dynamicVideo);
            Console.WriteLine("VerifyStream: streamID=" + streamId.ToString() +
                "; Start=" + startTime.ToString() + "; End=" + endTime.ToString());
            verifyStream(streamPlayer);
        }
        
        private static void verifyStream(DBStreamPlayer streamPlayer) {
            BufferChunk frame;
            long streamTime;
            BufferChunk sample;
            bool keyframe;
            long refTime = 0;
            long lastStreamTime = 0;
            long timediff = 0;
            DateTime dt;
            while (streamPlayer.GetNextFrame(out frame, out streamTime)) {
                try {
                    sample = ProfileUtility.FrameToSample(frame, out keyframe);
                }
                catch (Exception ex) {
                    dt = new DateTime(streamTime);
                    Console.WriteLine("FrameToSample failed at: " + dt.ToString() + "; Sampletime=" + streamTime.ToString());                    
                    Console.WriteLine(ex.ToString());
                    continue;
                }
                if (refTime == 0) 
                    refTime = streamTime;


                timediff = streamTime - lastStreamTime;

                // Look for large intervals
                //if (timediff > 500000L) {
                //    dt = new DateTime(streamTime);
                //    Console.WriteLine("Sample: " + dt.ToString() + "; Sampletime=" + streamTime.ToString() + "; length=" + sample.Length.ToString() + ";interval=" + timediff.ToString());
                //}
                // Look for large samples
                //if (sample.Length > 90000) {
                //    dt = new DateTime(streamTime);
                //    Console.WriteLine("Sample: " + dt.ToString() + "; Sampletime=" + streamTime.ToString() + "; length=" + sample.Length.ToString() + ";interval=" + timediff.ToString());
                //}
                // Look for small samples
                //if (sample.Length < 300) {
                //    dt = new DateTime(streamTime);
                //    Console.WriteLine("Sample: " + dt.ToString() + "; Sampletime=" + streamTime.ToString() + "; length=" + sample.Length.ToString() + ";interval=" + timediff.ToString());
                //}

                lastStreamTime = streamTime;
            }

            DateTime dt1 = new DateTime(refTime);
            DateTime dt2 = new DateTime(lastStreamTime);
            Console.WriteLine("Started at " + dt1.ToString() + "; Ended at " + dt2.ToString() + " (" + lastStreamTime.ToString() + " ticks)" + 
                " after duration ticks =" + (lastStreamTime - refTime).ToString());
        }



        internal static void UnPillarBoxTest() {
            Console.WriteLine("UnPillarbox Test");
            // Need a sample file to work with:
            string fn = @"c:\pbtest.jpg";
            string outfn = @"c:\unpillarboxed.jpg";
            if (!File.Exists(fn)) {
                Console.WriteLine(" ! Failed: Input file not found: " + fn);
                return;
            }
            try {
                Bitmap bm = (Bitmap)Bitmap.FromFile(fn);
                Bitmap fixedbm = PresentationFromVideoMgr.TestUnPillarBox(bm);
                fixedbm.Save(outfn);
            }
            catch (Exception ex) {
                Console.WriteLine(" ! Failed: " + ex.ToString());
                return;
            }
            Console.WriteLine(" Look at " + outfn + " to verify correct result.");
        }

        #endregion Test Code    
    
    }
}
