using System;
using System.Text;
using System.IO;
using System.Xml.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Diagnostics;
using MSR.LST;
using MSR.LST.Net.Rtp;
using System.Text.RegularExpressions;
using ArchiveRTNav;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Static utility methods.
	/// </summary>
	public class Utility
	{
		public Utility(){}

        /// <summary>
        /// If the segment has the specified flag set, return true.
        /// </summary>
        /// <param name="segment"></param>
        /// <param name="flag"></param>
        /// <returns></returns>
        public static bool SegmentFlagIsSet(ArchiveTranscoderJobSegment segment, SegmentFlags flag)
        {
            if (segment != null)
            {
                if (segment.Flags != null)
                {
                    foreach (string s in segment.Flags)
                    {
                        if (s != null)
                        {
                            if (Enum.GetName(typeof(SegmentFlags), flag).ToLower() == s.Trim().ToLower())
                                return true;
                        }
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Add the specified flag to segment flags.  If the flag is already set, do nothing.
        /// </summary>
        /// <param name="segment"></param>
        /// <param name="flag"></param>
        public static void SetSegmentFlag(ArchiveTranscoderJobSegment segment, SegmentFlags flag)
        {

            if (segment != null)
            {
                if (segment.Flags == null)
                {
                    segment.Flags = new String[1];
                    segment.Flags[0] = Enum.GetName(typeof(SegmentFlags), flag);
                }
                else
                {
                    if (SegmentFlagIsSet(segment, flag))
                        return;

                    String[] newFlags = new String[segment.Flags.Length + 1];
                    segment.Flags.CopyTo(newFlags, 0);
                    newFlags[newFlags.Length-1] = Enum.GetName(typeof(SegmentFlags), flag);
                    segment.Flags = newFlags;
                }
            }
        }


		/// <summary>
		/// Return true if string input contains "/" or "\";
		/// </summary>
		/// <param name="s"></param>
		/// <returns></returns>
		public static bool ContainsSlash(String s)
		{
			if ((s.IndexOf("/") != -1) || (s.IndexOf(@"\") != -1))
				return true;

			return false;
		}

        private static bool pptInstalled = false;
        private static bool checkedPptInstalled = false;

		/// <summary>
		/// Test to see if PPT is installed on this system.
        /// After we've checked once, just return the same value again without rechecking.
		/// </summary>
		/// <returns></returns>
		public static bool CheckPptIsInstalled()
		{
            if (checkedPptInstalled) {
                return pptInstalled;
            }

			bool ret = true;
			Microsoft.Office.Interop.PowerPoint.Application ppApp;
			try
			{
				ppApp = new Microsoft.Office.Interop.PowerPoint.Application();
			}
			catch
			{
				ret = false;
			}
			finally
			{
				ppApp = null;
			}

            checkedPptInstalled = true;
            pptInstalled = ret;
			return ret;
		}

		/// <summary>
		/// Serialize an object
		/// </summary>
		/// <param name="b"></param>
		/// <returns></returns>
		public static byte[] ObjectToByteArray(Object b)
		{
			if (b==null)
				return new byte[0];
			BinaryFormatter bf = new BinaryFormatter();
			MemoryStream ms = new MemoryStream();
			try
			{
				bf.Serialize(ms,b);
			}
			catch (Exception e)
			{
				Debug.WriteLine(e.ToString());
				return new byte[0];
			}
			ms.Position = 0;//rewind
			byte[] ba = new byte[ms.Length];
			ms.Read(ba,0,(int)ms.Length);
			return ba;
		}
	
		/// <summary>
		/// Deserialize an object
		/// </summary>
		/// <param name="ba"></param>
		/// <returns></returns>
		public static Object ByteArrayToObject(byte[] ba)
		{
			BinaryFormatter bf = new BinaryFormatter();
			MemoryStream ms = new MemoryStream(ba);
			ms.Position = 0;
			try
			{
				return (Object) bf.Deserialize(ms);
			}
			catch(Exception e)
			{
				Debug.WriteLine(e.ToString());
				return null;
			}
		}


		/// <summary>
		/// Generate a path to a new file in the specified temp area with the given extension.
		/// </summary>
		/// <param name="extent"></param>
		/// <returns></returns>
		public static String GetTempFilePath(String extent)
		{
			String path = Path.Combine(Constants.TempDirectory,"ArchiveTranscoderTemp");
			int index = 0;
			while (File.Exists(path + index.ToString() + "." + extent))
			{
				index++;
			}
			return path + index.ToString() + "." + extent;
		}

		/// <summary>
		/// Generate a path to a new directory in the specified temp area
        /// It is for the caller to create the directory.  This should be done promptly
        /// since another thread could also be assigned the same name until the directory exists.
		/// </summary>
		/// <param name="extent"></param>
		/// <returns></returns>
		public static String GetTempDir()
		{
			String path = Path.Combine(Constants.TempDirectory, "ArchiveTranscoderTemp");
			int index = 0;
			while (Directory.Exists(path + index.ToString()))
			{
				index++;
			}
			return path + index.ToString();
		}

		/// <summary>
		/// Return a list of ArchiveTranscoderTemp items.
		/// </summary>
		/// <returns></returns>
		public static String GetExistingTempFilesAndDirs()
		{
			String path = Constants.TempDirectory;
			String[] files = Directory.GetFiles(path,"ArchiveTranscoderTemp*");
			String[] dirs = Directory.GetDirectories(path,"ArchiveTranscoderTemp*");
			StringBuilder sb = new StringBuilder();
			foreach (String s in files)
			{
				sb.Append(s + "\r\n");
			}
			foreach (String s in dirs)
			{
				sb.Append(s + "\r\n");
			}
			return sb.ToString();
		}


		/// <summary>
		/// Delete ArchiveTranscoderTemp items
		/// </summary>
		public static void DeleteExistingTempFilesAndDirs()
		{
			String path = Constants.TempDirectory;
			String[] files = Directory.GetFiles(path,"ArchiveTranscoderTemp*");
			String[] dirs = Directory.GetDirectories(path,"ArchiveTranscoderTemp*");
			foreach (String s in files)
			{
				try
				{
					File.Delete(s);
				}
				catch
				{
					Debug.WriteLine("Failed to delete file: " + s);
				}
			}
			foreach (String s in dirs)
			{
				try
				{
					Directory.Delete(s,true);
				}
				catch
				{
					Debug.WriteLine("Failed to delete directory: " + s);
				}
			}
		}


		/// <summary>
		/// Deserialize XML Batch file.
		/// </summary>
		/// <param name="filename"></param>
		/// <returns></returns>
		/// Todo: add some error handling.
		public static ArchiveTranscoderBatch ReadBatchFile(string filename)
		{
			XmlSerializer serializer = new XmlSerializer(typeof(ArchiveTranscoderBatch));

			// If the XML document has been altered with unknown 
			// nodes or attributes, they could be handled with
			// UnknownNode and UnknownAttribute events.
			//serializer.UnknownNode+= new XmlNodeEventHandler(serializer_UnknownNode);
			//serializer.UnknownAttribute+= new XmlAttributeEventHandler(serializer_UnknownAttribute);  
			
			ArchiveTranscoderBatch b = null;
			FileStream fs = null;
			try
			{
				fs = new FileStream(filename, FileMode.Open,FileAccess.Read);
				b = (ArchiveTranscoderBatch) serializer.Deserialize(fs);
			}
			catch
			{
				b = null;
			}
			finally
			{
                if (fs != null)
				    fs.Close();
			}
			return b;
		}

		/// <summary>
		/// Write Batch file
		/// </summary>
		public static void WriteBatch(ArchiveTranscoderBatch batch, String filename)
		{
			XmlSerializer serializer = 
				new XmlSerializer(typeof(ArchiveTranscoderBatch));
			TextWriter writer = new StreamWriter(filename);
			serializer.Serialize(writer, batch);
			writer.Close();
		}

		static char[] hexDigits = { '0', '1', '2', '3', '4', '5', '6', '7',
									'8', '9', 'A', 'B', 'C', 'D', 'E', 'F'};

		/// <summary>
		/// Convert a byte[] to a string of hex
		/// </summary>
		/// <param name="ba"></param>
		/// <returns></returns>
		public static String ByteArrayToHexString(byte[] ba)
		{
			char[] chars = new char[ba.Length * 2];
			for (int i = 0; i < ba.Length; i++) 
			{
				int b = ba[i];
				chars[i * 2] = hexDigits[b >> 4];
				chars[i * 2 + 1] = hexDigits[b & 0xF];
			}
			return new string(chars);
		}

		/// <summary>
		/// Make a unique file name by adding or incrementing a numeral on the root part.  
		/// If the input names a file that does not exist, just return it.
		/// </summary>
		/// <param name="rawname">user-supplied input</param>
		/// <returns>unique file name</returns>
		public static string GenerateUniqueFileName(string rawname)
		{
			FileInfo fi = new FileInfo(rawname);
			if (!fi.Exists) 
				return rawname;

			string root, extent;
			int num;
			string left, right;

			if (rawname.LastIndexOf(".") <= 0)
			{
				root = rawname;
				extent = ".wmv";
			}
			else
			{
				root = rawname.Substring(0,rawname.LastIndexOf("."));
				extent = rawname.Substring(rawname.LastIndexOf("."));
			}

			Regex re = new Regex("^(.*?)([0-9]*)$");
			Match match = re.Match(root); 
			if (match.Success)
			{
				left = match.Result("$1");
				right = match.Result("$2");
				if (right != "")
					num = Convert.ToInt32(right);
				else
					num=0;
			}
			else
			{
				num = 0;
				left = root;
			}
			num++;

			while (true)
			{
				fi = new FileInfo(left + num.ToString() + extent);
				if (!fi.Exists)
					break;
				else
					num++;
			}

			return left + num.ToString() + extent;
		}

		/// <summary>
		/// Read a CSD and get its document guid
		/// </summary>
		/// <param name="fileName"></param>
		/// <returns></returns>
		public static Guid GetGuidFromCsd(String fileName)
		{
			Guid ret = Guid.Empty;
			SlideViewer.CSDDocument document = null;
			BinaryFormatter binaryFormatter = new BinaryFormatter();
			FileStream stream = null;
			try
			{
				stream = new FileStream(fileName,FileMode.Open,FileAccess.Read,FileShare.ReadWrite);
				document = (SlideViewer.CSDDocument)binaryFormatter.Deserialize(stream);
				ret = document.Identifier;
			}
			catch
			{
				Debug.WriteLine("Failed to read: " + fileName);
			}
			finally
			{
				stream.Close();
			}
			return ret;	
		}

		
		/// <summary>
		/// Test that a string can be sucessfully parsed as a URI.
		/// </summary>
		/// <param name="uri"></param>
		/// <returns></returns>
		public static bool isUri(String uri)
		{
			try
			{
				Uri u = new Uri(uri.Trim());
			}
			catch
			{
				return false;
			}
			return true;
		}

		/// <summary>
		/// Test that a string can be parsed as DateTime
		/// </summary>
		/// <param name="s"></param>
		/// <returns></returns>
		public static bool isDateTime(String s)
		{
			if (s==null)
				return false;
			try
			{
				DateTime foo = DateTime.Parse(s);
			}
			catch
			{
				return false;
			}
			return true;
		}

		/// <summary>
		/// Test that a string is non-null and contains at least one non-whitespace character.
		/// </summary>
		/// <param name="s"></param>
		/// <returns></returns>
		public static bool isSet(String s)
		{
			if (s==null)
				return false;
			if (s.Trim() == "")
				return false;
			return true;
		}	

		/// <summary>
		/// Convert a string to a PayloadType for any supported payload.  For unsupported payloads
		/// return xApplication2.
		/// </summary>
		/// <param name="s"></param>
		/// <returns></returns>
		public static PayloadType StringToPayloadType(String s)
		{

			PayloadType payload = PayloadType.xApplication2;
			switch (s)
			{
				case "dynamicAudio":
				{
					payload = PayloadType.dynamicAudio;
					break;
				}
				case "dynamicVideo":
				{
					payload = PayloadType.dynamicVideo;
					break;
				}
				case "dynamicPresentation":
				{
					payload = PayloadType.dynamicPresentation;
					break;
				}
				case "RTDocument":
				{
					payload = PayloadType.RTDocument;
					break;
				}
			}
			return payload;
		}

		/// <summary>
		/// Convert a string to a PresenterWireFormatType.
		/// </summary>
		/// <param name="s"></param>
		/// <returns></returns>
		public static PresenterWireFormatType StringToPresenterWireFormatType(String s)
		{
			foreach (PresenterWireFormatType t in Enum.GetValues(typeof(PresenterWireFormatType)))
			{
				if (s==t.ToString())
				{
					return t;
				}
			}
			return PresenterWireFormatType.Unknown;
		}

		/// <summary>
		/// Infer a payload type from a format
		/// </summary>
		/// <param name="f"></param>
		/// <returns></returns>
        public static PayloadType formatToPayload(PresenterWireFormatType f) {
            PayloadType p = PayloadType.xApplication2;
            switch (f) {
                case PresenterWireFormatType.CPNav: {
                        p = PayloadType.dynamicPresentation;
                        break;
                    }
                case PresenterWireFormatType.CPCapability: {
                        p = PayloadType.RTDocument;
                        break;
                    }
                case PresenterWireFormatType.CP3: {
                        p = PayloadType.dynamicPresentation;
                        break;
                    }
                case PresenterWireFormatType.RTDocument: {
                        p = PayloadType.RTDocument;
                        break;
                    }
                case PresenterWireFormatType.Video: {
                        p = PayloadType.dynamicVideo;
                        break;
                    }
            }
            return p;
        }

		/// <summary>
		/// Return the int32 at the given position in a byte[].
		/// </summary>
		/// <param name="buffer"></param>
		/// <param name="index"></param>
		/// <param name="littleEndian"></param>
		/// <returns></returns>
		/// From MSR's BufferChunk
		public static Int32 GetInt32(byte[] buffer, int index, bool littleEndian)
		{
			Int32 ret;
            
			// BigEndian network -> LittleEndian architecture
			if(littleEndian)
			{
				ret  = buffer[index + 0] << 3 * 8;
				ret += buffer[index + 1] << 2 * 8;
				ret += buffer[index + 2] << 1 * 8;
				ret += buffer[index + 3] << 0 * 8;
			}
			else // BigEndian network -> BigEndian architecture
			{
				ret  = buffer[index + 0] << 0 * 8;
				ret += buffer[index + 1] << 1 * 8;
				ret += buffer[index + 2] << 2 * 8;
				ret += buffer[index + 3] << 3 * 8;
			}

			return ret;
		}

        public static string GetLocalizedDateTimeString(DateTime dt, string timeformat) {
            return dt.ToString("d") + " " + dt.ToString(timeformat);
        }
	}
}
