using System;
using System.Diagnostics;
using MSR.LST;
using MSR.LST.MDShow;
using MSR.LST.Net.Rtp;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Encapsulate buffer, pointers and database access operations for a stream_id and time range.
	/// maintiain a buffer and array of indices, refilling both from the database as necessary.
	/// GetNextFrame method returns the byte[] and timestamp for the next frame found for the stream.
	/// 
	/// Reference: MSR Archiver BufferPlayer.cs
	/// </summary>
	public class DBStreamPlayer
	{
		#region Members

		private int streamID;          // One and only DB streamID we will ever use.
		private long start;            // Don't deliver stream data before this time 
		private long end;              // Don't deliver stream data after this time
		private int currentIndex;      // pointer to next unread element in indices
		private int indexCount;        // Current value of indices.Length
        private int minOffset;         // The start offset of the currently loaded buffer
        private int maxOffset;         // The end offset of the currently loaded buffer
		private long startingTick;     // time of the current first element in indices
		private long endingTick;       // time of the current last element in indices
		private bool streamEndReached; // All the stream data before this.end has been exhausted.
		private BufferChunk frame;	   // buffer
		private Index[] indices;       // pointers to frames within the current buffer
		private MediaType mediaType;   // The MediaType struct for this stream (if audio or video)
		private long firstTick;		   // The time of the first frame for this stream in the database.

		#endregion Members

		#region Constructors

		/// <summary>
		/// Read a specified time range in the stream
		/// </summary>
		public DBStreamPlayer(int streamID, long start, long end, PayloadType payload)
		{
			this.streamID = streamID;
			this.start = start;
			this.end = end;
			currentIndex = 0;
			indexCount = 0;
			startingTick = start;
			endingTick = start-1;
			streamEndReached = false;

			//Get Stream parameters
			int maxFrameSize;
			int maxFrameCount;
			int maxBufferSize;
			DatabaseUtility.GetStreamStatistics( streamID, out firstTick, out maxFrameSize, out maxFrameCount, out maxBufferSize);
			
			//Allocate buffer and indices array
			frame = new BufferChunk( maxBufferSize );
			indices = new Index[maxFrameCount]; 

			//Set the stream MediaType
			mediaType = null;
			if (payload == PayloadType.dynamicVideo)
			{
				mediaType = ProfileUtility.StreamIdToVideoMediaType(streamID);
			}
			else if (payload == PayloadType.dynamicAudio)
			{
				mediaType = ProfileUtility.StreamIdToAudioMediaType(streamID);
			}
		}

		/// <summary>
		/// Read 'duration' ticks from the start of the stream.  Note: payload param is only significant for audio and video
		/// </summary>
		public DBStreamPlayer(int streamID, long duration, PayloadType payload)
		{
			this.streamID = streamID;
			currentIndex = 0;
			indexCount = 0;
			streamEndReached = false;

			//Get Stream parameters
			int maxFrameSize;
			int maxFrameCount;
			int maxBufferSize;
			DatabaseUtility.GetStreamStatistics( streamID, out firstTick, out maxFrameSize, out maxFrameCount, out maxBufferSize);

			//Allocate buffer and indices array
			frame = new BufferChunk( maxBufferSize );
			indices = new Index[maxFrameCount]; 

			//Init start and end
			this.start = firstTick;
			this.end = start + duration;
			startingTick = start;
			endingTick = start-1;
			
			//Set the stream MediaType
			mediaType = null;
			if (payload == PayloadType.dynamicVideo)
			{
				mediaType = ProfileUtility.StreamIdToVideoMediaType(streamID);
			}
			else if (payload == PayloadType.dynamicAudio)
			{
				mediaType = ProfileUtility.StreamIdToAudioMediaType(streamID);
			}
		}

		#endregion Constructors

		#region Properties

		public DateTime Start
		{
			get { return new DateTime(start); }
		}

		public DateTime End
		{
			get { return new DateTime(end); }
		}

		public MediaType StreamMediaType
		{
			get { return mediaType; }
		}

		/// <summary>
		/// The time in ticks of the first frame of this stream in the database 
		/// </summary>
		public long FirstTick
		{
			get { return this.firstTick; }
		}

		#endregion Properties

		#region Public Methods

		/// <summary>
		/// retrieve the next bufferchunk and timestamp for the stream.  Return false if there are no more.
		/// </summary>
		public bool GetNextFrame(out BufferChunk outFrame, out long timestamp)
		{
			outFrame = null;
			timestamp = 0;
			bool ret = false;

			if (streamEndReached)
			{
				return false;
			}

			if ( currentIndex >= indexCount)
			{
				FillBuffer();
			}

			if (currentIndex < indexCount)
			{
				int frameLength = 1 + indices[currentIndex].end - indices[currentIndex].start;

				if ( frameLength > 0)
				{
					if (indices[currentIndex].timestamp < end)
					{
						frame.Reset( indices[currentIndex].start - this.minOffset, frameLength );
						outFrame = frame;
						timestamp = indices[currentIndex].timestamp;
						ret = true;
					}
					else
					{
						streamEndReached = true;
						ret = false;
					}
				}
				else
				{
					Debug.Fail("Frame of length zero found.");
					ret = false;
				}

				++currentIndex;			
				return ret;
			}
			else
			{
				return false;
			}
		}


		/// <summary>
		/// Get the timestamp of the frame to be returned on the next call to GetNextFrame.  
		/// Return false if there are no more frames.
		/// </summary>
		/// <param name="timestamp"></param>
		/// <returns></returns>
		public bool GetNextFrameTime(out long timestamp)
		{
			timestamp = 0;
			bool ret = false;

			if (streamEndReached)
			{
				return false;
			}

			if ( currentIndex >= indexCount)
			{
				FillBuffer();
			}

			if (currentIndex < indexCount)
			{
				int frameLength = 1 + indices[currentIndex].end - indices[currentIndex].start;

				if ( frameLength > 0)
				{
					if (indices[currentIndex].timestamp < end)
					{
						timestamp = indices[currentIndex].timestamp;
						ret = true;
					}
					else
					{
						streamEndReached = true;
						ret = false;
					}
				}
				else
				{
					Debug.Fail("Frame of length zero found.");
					ret = false;
				}

				return ret;
			}
			else
			{
				return false;
			}		
		}

		#endregion Public Methods

		#region Private Methods

		private void FillBuffer()
		{
			// load the indices from the database
			indexCount = 0;
			currentIndex = 0;
            bool continuousData;

		    bool streamOutOfData = DatabaseUtility.LoadIndices(indices, endingTick+1, streamID, frame.Buffer.Length, 
                out indexCount, out minOffset, out maxOffset, out continuousData);

            //Debug.WriteLine("FillBuffer for StreamID=" + streamID.ToString() + "; starting tick=" + (endingTick + 1).ToString());
            
            // load the associated buffer from the database if we're not empty....
			if ( indexCount > 0 ) {
                /// Note on data quirks: Sometimes the raw data from the archive service is out of order.
                /// I think this is a bug in Archive Service.  The indices are in order by timestamp, but the start/end fields may not
                /// align one after another.  In this case there are also corresponding discontinuities in the raw data.
                /// To work around this we have LoadIndices keep track of the minimum start and the maximum end in the
                /// indices we've just fetched.  They are not necessarily the first and last items in the array.
                /// We also add a check to make sure we don't try to fetch more data than the frame buffer can
                /// hold.  

                if (!continuousData) {
                    DateTime dt = new DateTime(this.endingTick + 1);
                    Console.WriteLine("DBStreamPlayer.FillBuffer: A data discontinuity was found for streamID=" + streamID.ToString() + ";near time=" + dt.ToString());
                }

				this.startingTick = indices[0].timestamp;
				this.endingTick = indices[indexCount-1].timestamp;

                DatabaseUtility.LoadBuffer(streamID, minOffset, maxOffset, ref frame);
			}
			else
			{
				streamEndReached = true;
			}
		}

        /// <summary>
        /// Manually go through an array of Index and return the minimum start and maximum end.
        /// Return false if min or max were other than the first.start and last.end.
        /// </summary>
        /// <param name="ind"></param>
        /// <param name="max"></param>
        /// <param name="min"></param>
        private bool findMaxMin(Index[] ind, out int max, out int min) {
            min = 0;
            max = 0;
            bool ret = true;
            if (ind.Length > 0) {
                min = ind[0].start;
                max = ind[ind.Length - 1].end;
                foreach (Index i in indices) {
                    if (i.start < min) {
                        min = i.start;
                        ret = false;
                    }
                    if (i.end > max) {
                        max = i.end;
                        ret = false;
                    }
                }
            }
            return ret;
        }

		#endregion Private Methods
	}
}
