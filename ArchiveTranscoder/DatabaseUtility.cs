using System;
using System.Diagnostics;
using System.Collections;
using System.Data.SqlClient;
using System.Data;
using MSR.LST.Net.Rtp;
using MSR.LST;

namespace ArchiveTranscoder
{
	/// <summary>
	/// All the static methods and the connection string used for database access.  Some of the methods come more or
	/// less directly from the Archive Service source.
	/// </summary>
	public class DatabaseUtility
	{
		public static String SQLConnectionString = "";

		public DatabaseUtility() {}

		#region Private

		private static SqlConnection conn;
		private static SqlCommand cmd;
		private static SqlDataReader reader;

		/// <summary>
		/// Create a new SqlConnection and SqlCommand of type Text, and execute the SqlDataReader.
		/// Clean up and return false if any problems are encountered.
		/// </summary>
		/// <param name="query"></param>
		/// <returns></returns>
		private static bool openAndExecute(String query)
		{
			bool retval = true;
			conn = new SqlConnection(SQLConnectionString);
			cmd = new SqlCommand(query, conn);
			cmd.CommandType = CommandType.Text;

			try 
			{
				conn.Open();
			}
			catch (Exception e)
			{
				Debug.WriteLine("Failed to connect to SQL Server.  Verify server is active and connection string is correct. Exception text: \n" + e.ToString());
				retval = false;
			}

			if (retval)
			{
				try
				{
					reader = cmd.ExecuteReader();
				}
				catch (Exception e)
				{
					Debug.WriteLine("Failed to find expected SQL schema.  Verify schema is installed correctly. Exception text: \n" + e.ToString());
					retval = false;
				}
				if (!retval)
				{
					try
					{
						reader.Close();
					}
					catch {}
				}
			}
			
			if (!retval)
			{
				try
				{
					conn.Close();
				}
				catch {}
			}

			return retval;
		}

		/// <summary>
		/// Close SqlDataReader and SqlConnection.
		/// </summary>
		private static void cleanUpConnection()
		{
			if (reader != null)
			{
				try
				{reader.Close();}
				catch
				{}
			}

			if (conn != null)
			{
				try
				{conn.Close();}
				catch
				{}
			}
		}


		/// <summary>
		/// Submit a query to return a single Int32
		/// </summary>
		/// <returns></returns>
		private static int numericQuery(String query)
		{
			if (!DatabaseUtility.openAndExecute(query))
			{
				return -1;
			}

			int count=-1;

			try 
			{
				if (reader.Read()) 
				{
					count = reader.GetInt32(0);
				}
			}
			finally 
			{
				DatabaseUtility.cleanUpConnection();
			}

			return count;
		}

		#endregion Private

		#region Public Static Methods

		/// <summary>
		/// Verify basic connectivity using currently set server, db name, and custom connection string, if any.  
		/// No checks are made here on the correctness of the schema.
		/// </summary>
		/// <returns></returns>
		public static bool CheckConnectivity()
		{
			// The timeout specified does not seem to be observed.
			// PRI2: It would be great to do this check faster.  Maybe do a quick dns resolve and ping the sql port before
			// trying to open the connection?  Maybe .Net has a better way to do this in one of the System.Data APIs?
			SqlConnection conn = null;
			bool retval = true;
			try
			{
				conn = new SqlConnection("Timeout=3;" + SQLConnectionString);
				conn.Open();
			}
			catch
			{
				retval = false;
			}
			finally
			{
				conn.Close();
			}

			return retval;
		}

		public static long GetConferenceStartTime(PayloadType payload, String cname, long start, long end)
		{
			String query = "select c.start_dttm " +
				"from conference c " +
				"join participant p on c.conference_id=p.conference_id " +
				"join stream s on p.participant_id=s.participant_id " +
				"join frame f on s.stream_id=f.stream_id " +
				"where s.payload_type='" + payload.ToString() + "' " +
				"and p.cname='" + cname + "' " +
				"and f.frame_time > " + start.ToString() + " " +
				"and f.frame_time < " + end.ToString() + " " +
				"group by c.start_dttm";
			//Debug.WriteLine(query);

			if (!DatabaseUtility.openAndExecute(query))
			{
				return 0;
			}

			long ret = 0;
			try 
			{
				if (reader.Read())
				{
					DateTime dt = reader.GetDateTime(0);
					ret = dt.Ticks;
				}
			}
			finally 
			{
				DatabaseUtility.cleanUpConnection();
			}

			return ret;

		}


		/// <summary>
		/// For the given cname, stream name, payload and time range, return an array of relevant stream_id's
		/// The array is in order by earliest frame time.  If the stream name is null it will be excluded
		/// from the where clause.
		/// </summary>
		public static int[] GetStreams(PayloadType payload, String cname, String name, long start, long end)
		{
			String query = "select s.stream_id, min(f.frame_time) as st_start " +
					"from participant p " +
					"join stream s on p.participant_id=s.participant_id " +
					"join frame f on s.stream_id=f.stream_id " +
					"where s.payload_type='" + payload.ToString() + "' " +
					"and p.cname='" + cname + "' ";
			if ((name!=null) && (name.Trim() != ""))
			{
				query += "and s.name='" + name + "' ";
			}
			query += "and f.frame_time > " + start.ToString() + " " +
					"and f.frame_time < " + end.ToString() + " " +
					"group by s.stream_id " +
					"order by st_start";

			if (!DatabaseUtility.openAndExecute(query))
			{
				return null;
			}

			ArrayList al = new ArrayList();
			try 
			{
				while (reader.Read()) 
				{
					al.Add(reader.GetInt32(0));
				}
			}
			finally 
			{
				DatabaseUtility.cleanUpConnection();
			}

			if (al.Count != 0)
			{
				return (int[])al.ToArray(typeof(int));
			}
			return null;
		}

		
		/// <summary>
		/// Return the array of streams for the given participant ID using the standard ArchiveService SP.
		/// </summary>
		/// <param name="participantID"></param>
		/// <returns></returns>
		static public Stream[] GetStreams( int participantID )
		{
			SqlConnection conn = new SqlConnection(SQLConnectionString);
			try
			{
				SqlCommand cmd = new SqlCommand("GetStreams", conn);
				cmd.CommandType = CommandType.StoredProcedure;

				SqlParameter sqlParticipantID = cmd.Parameters.Add("@participant_id", SqlDbType.Int);
				sqlParticipantID.Direction = ParameterDirection.Input;
				sqlParticipantID.Value = participantID;

				conn.Open();
				SqlDataReader r = cmd.ExecuteReader(CommandBehavior.SingleResult);

				ArrayList streamList = new ArrayList(10);
				while( r.Read())
				{
					Stream stream= new Stream(  
						r.GetInt32(0),  //stream id
						r.GetString(1), // name
						r.GetString(2), // payload,
						r.GetInt32(3),  // frames
						r.IsDBNull(4) ? 0L : r.GetInt64(4), // seconds
						r.GetInt32(5)); // bytes 
					streamList.Add( stream);
				}

				r.Close();
				conn.Close();

				return ( Stream[]) streamList.ToArray( typeof(Stream));
			}
			catch( SqlException ex )
			{
				Debug.WriteLine( "Database operation failed.  Exception: \n" + ex.ToString());
			}
			finally
			{
				conn.Close();
			}
			return null;
		}

		/// <summary>
		/// Return a count of rows in the participant table.  -1 indicates an error.
		/// </summary>
		/// <returns></returns>
		public static int CountParticipants()
		{
			return numericQuery("select count(*) from participant");
		}

		/// <summary>
		/// Return the first frame for the given stream.
		/// </summary>
		/// <param name="streamID"></param>
		/// <returns></returns>
		public static byte[] GetFirstFrame(int streamID)
		{
			String query = "select top 1 f.raw_start, f.raw_end " +
				"from stream s join frame f on s.stream_id=f.stream_id " +
				"where s.stream_id=" + streamID.ToString() + " " +
				"order by f.frame_time ";

			if (!DatabaseUtility.openAndExecute(query))
			{
				return null;
			}

			int raw_start=-1;
			int raw_end=-1;

			try 
			{
				if (reader.Read()) 
				{
					raw_start=reader.GetInt32(0);
					raw_end=reader.GetInt32(1);
				}
			}
			finally 
			{
				DatabaseUtility.cleanUpConnection();
			}

			if (raw_start!=-1)
			{
				BufferChunk ret = new BufferChunk(raw_end - raw_start + 1);
				LoadBuffer(streamID, raw_start, raw_end, ref ret);
				return ret.Buffer;
			}
			return null;
		}

		/// <summary>
		/// Get the first frame for payload/cname after start and before end.  Return null
		/// if such frame does not exist in the database.
		/// </summary>
		/// <param name="payload"></param>
		/// <param name="cname"></param>
		/// <param name="start"></param>
		/// <param name="end"></param>
		/// <returns></returns>
		public static byte[] GetFirstFrame(PayloadType payload, String cname, String name, long start, long end)
		{
			// get the indices and stream_id:
			//DateTime startDt = DateTime.Parse(start);
			//DateTime endDt = DateTime.Parse(end);
			String query = "select top 1 s.stream_id, f.frame_id , f.raw_start, f.raw_end " +
					"from participant p join stream s on p.participant_id=s.participant_id " +
					"join frame f on s.stream_id=f.stream_id where s.payload_type='" + 
					payload.ToString() + "' and p.cname='" + cname + "' ";

			if ((name != null) && (name.Trim() != ""))
			{
				query += "and s.name='" + name + "' ";
			}
			query += "and f.frame_time >= " + start.ToString() + " and f.frame_time < " +
					end.ToString() + " order by frame_time ";

			if (!DatabaseUtility.openAndExecute(query))
			{
				return null;
			}

			int stream_id=-1;
			int frame_id=-1;
			int raw_start=-1;
			int raw_end=-1;

			try 
			{
				if (reader.Read()) 
				{
					stream_id = reader.GetInt32(0);
					frame_id=reader.GetInt32(1);
					raw_start=reader.GetInt32(2);
					raw_end=reader.GetInt32(3);
				}
			}
			finally 
			{
				DatabaseUtility.cleanUpConnection();
			}

			//Debug.WriteLine("stream_id="+stream_id.ToString());
			if (stream_id!=-1)
			{
				BufferChunk ret = new BufferChunk(raw_end - raw_start + 1);
				LoadBuffer(stream_id, raw_start, raw_end, ref ret);
				return ret.Buffer;
			}
			return null;
		}


		/// <summary>
		/// Return true if the database contains no data for the given payload/cname/name/timerange.
		/// if name is null or empty, exclude it from the where clause.
		/// </summary>
		public static bool isEmptyStream(PayloadType payload, String cname, String name, long start, long end)
		{
			String query = "select count(*) " +
				"from participant p join stream s on p.participant_id=s.participant_id " +
				"join frame f on s.stream_id=f.stream_id where s.payload_type='" + 
				payload.ToString() + "' and p.cname='" + cname + "' ";

			if ((name != null) && (name.Trim() != ""))
				query += "and s.name='" + name + "' ";

			query += "and f.frame_time >= " + start.ToString() + " and f.frame_time < " +
				end.ToString();

			int count = DatabaseUtility.numericQuery(query);

			if (count==0)
				return true;
			return false;
		}


		/// <summary>
		/// Loads a block of bytes for a given stream.
		/// </summary>
		public static void LoadBuffer(int streamID, int start, int end, ref BufferChunk frame)
		{
			SqlConnection conn = new SqlConnection(SQLConnectionString);
			SqlCommand cmd = new SqlCommand( "ReadBuffer ", conn);
			cmd.CommandType = CommandType.StoredProcedure;

			SqlParameter sqlStreamID = cmd.Parameters.Add( "@stream_id", SqlDbType.Int);
			sqlStreamID.Direction = ParameterDirection.Input;
			sqlStreamID.Value = streamID;

			SqlParameter startParam = cmd.Parameters.Add("@start", SqlDbType.Int);
			startParam.Direction = ParameterDirection.Input;
			startParam.Value = start;

			int dataLength = end - start + 1;

			SqlParameter finish = cmd.Parameters.Add("@length", SqlDbType.BigInt);
			finish.Direction = ParameterDirection.Input;
			finish.Value = dataLength;

			conn.Open();
			SqlDataReader dr = cmd.ExecuteReader(CommandBehavior.SingleRow);

			if ( dr.Read())
			{
				Debug.Assert( dataLength <= frame.Buffer.Length );
				dr.GetBytes(0, 0, frame.Buffer, 0, dataLength);
			}

			dr.Close();
			conn.Close();
		}


		/// <summary>
		/// Loads the next block of indexes into an array of indices.
		/// </summary>
		internal static bool LoadIndices(Index[] indices, long startingTick, int streamID, int maxBytes, out int indexCount,
            out int minOffset, out int maxOffset, out bool continuousData)
		{
			SqlConnection conn = new SqlConnection(SQLConnectionString);
			SqlCommand cmd = new SqlCommand("LoadIndices", conn);
			cmd.CommandType = CommandType.StoredProcedure;

			SqlParameter stream_id = cmd.Parameters.Add("@stream_id", SqlDbType.Int);
			stream_id.Direction = ParameterDirection.Input;
			stream_id.Value = streamID;

			SqlParameter starting_tick = cmd.Parameters.Add("@starting_tick", SqlDbType.BigInt);
			starting_tick.Direction = ParameterDirection.Input;
			starting_tick.Value = startingTick;

			SqlParameter ending_tick = cmd.Parameters.Add("@count", SqlDbType.BigInt);
			ending_tick.Direction = ParameterDirection.Input;
			ending_tick.Value = indices.Length;

			conn.Open();
			SqlDataReader dr = cmd.ExecuteReader(CommandBehavior.SequentialAccess);

			indexCount = 0;
            minOffset = int.MaxValue;
            maxOffset = 0;
            bool streamOutOfData = true;
            continuousData = true;

            while (dr.Read()) {
                indices[indexCount].start = dr.GetInt32(0);
                indices[indexCount].end = dr.GetInt32(1);
                indices[indexCount].timestamp = dr.GetInt64(2);

                int provisionalMin = minOffset;
                int provisionalMax = maxOffset;
                if (indices[indexCount].start < minOffset) {
                    provisionalMin = indices[indexCount].start;
                }
                if (indices[indexCount].end > maxOffset) {
                    provisionalMax = indices[indexCount].end;
                }

                if ((indexCount != 0) && (indices[indexCount].start != indices[indexCount - 1].end + 1)) {
                    continuousData = false;
                }

                // Check to see if this would overflow our buffer.  Add one since we load the data at both endpoints.
                if ((provisionalMax - provisionalMin + 1) > maxBytes) {
                    streamOutOfData = false;
                    //This is one frame too many... Skip it.
                    break;
                }

                // Accept this frame.
                minOffset = provisionalMin;
                maxOffset = provisionalMax;
                indexCount++;

                // Make sure indices isn't full
                if (indexCount >= indices.Length) {
                    streamOutOfData = false;
                    break;
                }
            }

			// Close our connections
			dr.Close();
			conn.Close();

			return streamOutOfData;
		}

		/// <summary>
		/// Gets statistics on a stream for playback.
		/// </summary>
		static internal void GetStreamStatistics( int streamID, out long firstTick, out int maxFrameSize, out int maxFrameCount, out int maxBufferSize)
		{
			SqlConnection conn = new SqlConnection(SQLConnectionString);

			try 
			{
				SqlCommand cmd = new SqlCommand( "GetStreamStatistics ", conn);
				cmd.CommandType = CommandType.StoredProcedure;

				SqlParameter strID = cmd.Parameters.Add( "@stream_id", SqlDbType.Int);
				strID.Direction = ParameterDirection.Input;
				strID.Value = streamID;

				SqlParameter interval = cmd.Parameters.Add( "@interval", SqlDbType.Int);
				interval.Direction = ParameterDirection.Input;
				interval.Value = Constants.PlaybackBufferInterval;

				SqlParameter startTick = cmd.Parameters.Add("@starting_tick", SqlDbType.BigInt);
				startTick.Direction = ParameterDirection.Output;

				SqlParameter max_frame_size = cmd.Parameters.Add("@max_frame_size", SqlDbType.Int);
				max_frame_size.Direction = ParameterDirection.Output;

				SqlParameter max_frame_count = cmd.Parameters.Add("@max_frame_count", SqlDbType.Int);
				max_frame_count.Direction = ParameterDirection.Output;

				SqlParameter max_buffer_size = cmd.Parameters.Add("@max_buffer_size", SqlDbType.Int);
				max_buffer_size.Direction = ParameterDirection.Output;

				conn.Open();
				cmd.ExecuteNonQuery();

				firstTick = (long) startTick.Value;
				maxFrameSize = (int) max_frame_size.Value;
				maxFrameCount = (int) max_frame_count.Value;
				maxBufferSize = (int) max_buffer_size.Value;
			}
			catch (SqlException ex)
			{
				Debug.WriteLine( "Database operation failed.  Exception: \n" + ex.ToString());
				throw;
			}
			finally
			{
				conn.Close();
			}
		}

		static public Conference[] GetConferences()
		{
			SqlConnection conn = new SqlConnection( SQLConnectionString);
			try
			{
				SqlCommand cmd = new SqlCommand("GetConferences", conn);
				cmd.CommandType = CommandType.StoredProcedure;

				conn.Open();
				SqlDataReader r = cmd.ExecuteReader(CommandBehavior.SingleResult);

				ArrayList conList = new ArrayList(10);
				while( r.Read())
				{
					Conference conf = new Conference(   
						r.GetInt32(0),      // conference id
						r.GetString(1),     // conference description
						r.GetString(2),     // venue name
						r.GetDateTime(3),   // start date time
						r.IsDBNull(4) ? DateTime.MinValue : r.GetDateTime(4) ); // end date time
					conList.Add(conf);
				}

				r.Close();
				conn.Close();

				return ( Conference[]) conList.ToArray( typeof(Conference));
			}
			catch( SqlException ex )
			{
				Debug.WriteLine( "Database operation failed.  Exception: \n" + ex.ToString());
			}
			finally
			{
				conn.Close();
			}
			return null;
		}

		/// <summary>
		/// Returns the participants in the given conference.
		/// </summary>
		/// <param name="conferenceID"></param>
		/// <returns></returns>
		static public Participant[] GetParticipants(int conferenceID)
		{
			SqlConnection conn = new SqlConnection( SQLConnectionString);
			try
			{
				SqlCommand cmd = new SqlCommand("GetParticipants", conn);
				cmd.CommandType = CommandType.StoredProcedure;

				SqlParameter conference_id = cmd.Parameters.Add("@conference_id", SqlDbType.Int);
				conference_id.Direction = ParameterDirection.Input;
				conference_id.Value = conferenceID;

				conn.Open();
				SqlDataReader r = cmd.ExecuteReader(CommandBehavior.SingleResult);

				ArrayList partList = new ArrayList(10);
				while( r.Read())
				{
					Participant part = new Participant( 
						r.GetInt32(0),  //session id
						r.GetString(2), // cname
						r.GetString(3));// name
					partList.Add(part);
				}

				r.Close();
				conn.Close();

				return (Participant[]) partList.ToArray(typeof(Participant));
			}
			catch( SqlException ex )
			{
				Debug.WriteLine( "Database operation failed.  Exception: \n" + ex.ToString());
			}
			finally
			{
				conn.Close();
			}
			return null;
		}

		#endregion Public Static Methods

		#region Test

		/// <summary>
		/// Return stream private extenstions, just for testing..
		/// </summary>
		/// <param name="streamID"></param>
		/// <returns></returns>
		public static Hashtable GetStreamPrivateExtensions(int streamID)
		{
			String query = "select privextns from stream where stream_id=" + streamID;

			if (!DatabaseUtility.openAndExecute(query))
				return null;

			try 
			{
				if (reader.Read()) 
				{
					System.Data.SqlTypes.SqlBinary sb = reader.GetSqlBinary(0);
					byte[] ba = sb.Value;
					object o = Utility.ByteArrayToObject(ba);
					Hashtable h = (Hashtable)o;
					return h;
				}
			}
			finally 
			{
				DatabaseUtility.cleanUpConnection();
			}

			return null;
		}		
		
		
		/// <summary>
		/// A variation of GetStreams that also returns the start time for each stream.  Unused for now.
		/// </summary>
		/// <param name="participantID"></param>
		/// <returns></returns>
		public static Stream[] GetStreams2(int participantID )
		{
			String query = "SELECT	s.stream_id, s.name, s.payload_type, " +
				"( SELECT MIN (frame_time) FROM frame WHERE stream_id = s.stream_id ) as st_time, " +
				"( SELECT COUNT( frame_id) FROM frame WHERE stream_id = s.stream_id ) as frames, " +
				"( SELECT (MAX( frame_time ) - MIN(frame_time))/10000000 FROM frame WHERE stream_id = s.stream_id ) as seconds, " +
				"( SELECT datalength(data) FROM rawStream WHERE stream_id =  s.stream_id) as bytes " +
				"FROM stream as s " +
				"WHERE s.participant_id = " + participantID.ToString() + " AND  " +
				"(SELECT datalength(data) FROM rawStream WHERE stream_id =  s.stream_id) > 1";
			
			if (!DatabaseUtility.openAndExecute(query))
			{
				return null;
			}

			ArrayList al = new ArrayList();
			try 
			{
				while (reader.Read()) 
				{
					al.Add(new Stream(
						reader.GetInt32(0),  //stream id
						reader.GetString(1), // name
						reader.GetString(2), // payload,
						reader.GetInt64(3), //start time
						reader.GetInt32(4),  // frames
						reader.IsDBNull(5) ? 0L : reader.GetInt64(5), // seconds
						reader.GetInt32(6))); // bytes 

				}
			}
			finally 
			{
				DatabaseUtility.cleanUpConnection();
			}

			if (al.Count != 0)
			{
				return (Stream[])al.ToArray(typeof(Stream));
			}
			return null;

		}

		#endregion Test
	}

	#region Conference Class

	/// <summary>
	/// Represent a row from the conference table
	/// </summary>
	[Serializable]
	public class Conference
	{
		public Conference( int conferenceID, string description, string venueIdentifier, DateTime start, DateTime end)
		{
			this.ConferenceID = conferenceID;
			this.Description = description;
			this.VenueIdentifier = venueIdentifier;
			this.Start = start;
			this.End = end;
		}
		public Conference() {}
		public int      ConferenceID;
		public string   Description;
		public string   VenueIdentifier;
		public DateTime Start;
		public DateTime End;

		public override string ToString()
		{
			return  "ID: " + ConferenceID +
				"\nDescription: " + Description + 
				"\nVenueIdentifier: " + VenueIdentifier + 
				"\nStart time: " + Start.ToLongTimeString() +
				"\nEnd time: " + End.ToLongTimeString();
		}

	}

	#endregion Conference Class

	#region Participant Class

	/// <summary>
	/// Represent a row from the participant table
	/// </summary>
	[Serializable]
	public class Participant
	{
		public Participant(int participant, string cname, string name )
		{
			this.ParticipantID = participant;
			this.CName = cname;
			this.Name = name;
		}

		public Participant() {}
    
		public int      ParticipantID;
		public string   CName;
		public string   Name;

		public override string ToString()
		{
			return  "ID: " + ParticipantID +
				"\nCName: " + CName +
				"\nName: " + Name;
		}

	}

	#endregion Participant Class

	#region Stream Class

	/// <summary>
	/// Represent a row from the stream table
	/// </summary>
	[Serializable]
	public class Stream
	{
		public Stream( int streamID, string name, string payload, long start_time, int frames, long seconds, int bytes)
		{
			this.StartTime = start_time;
			this.StreamID = streamID;
			this.Name = name;
			this.Payload = payload;
			this.Frames = frames;
			this.Seconds = seconds;
			this.Bytes = bytes;	
		}

		public Stream( int streamID, string name, string payload, int frames, long seconds, int bytes)
		{
			this.StartTime=0;
			this.StreamID = streamID;
			this.Name = name;
			this.Payload = payload;
			this.Frames = frames;
			this.Seconds = seconds;
			this.Bytes = bytes;
		}
		public Stream() {}
		public int      StreamID;
		public string   Name;
		public string   Payload;
		public int      Frames;
		public long     Seconds;
		public long		StartTime;
		public int      Bytes;

		public override string ToString()
		{
			return  "ID: " + StreamID +
				"\nName: " + Name +
				"\nPayloadType: " + Payload +
				"\nFrames: " + Frames +
				"\nSeconds: " + Seconds +
				"\nBytes: " + Bytes;
		}

	}

	#endregion Stream Class

	#region Index Struct

	/// <summary>
	/// Represents an index into a buffer of frames
	/// </summary>
	public struct Index
	{
		public int start;                   // start offest
		public int  end;                    // end offset
		public long timestamp;              // time in ticks
	}
	#endregion Index Struct
}

