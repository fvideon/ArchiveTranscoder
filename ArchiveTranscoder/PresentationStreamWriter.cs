using System;
using System.IO;
using System.Diagnostics;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Wrap a StreamWriter and create temp files, automatically inserting header and footer. 
	/// Support writing different versions of the file with different headers.  
	/// Typical usage sequence: construct, Write ..., Close, WriteHeaderAndCopy ..., release reference.
	/// </summary>
	public class PresentationStreamWriter
	{
		private StreamWriter streamWriter;
		private string fileBody;
		private LogMgr log;

		public String FileBody
		{
			get {return fileBody;}
		}

		#region Ctor/Dtor

		public PresentationStreamWriter(LogMgr log)
		{
			fileBody = Utility.GetTempFilePath("xml");
			this.log = log;
   			FileStream fileStream = new FileStream(fileBody,
				FileMode.Append, FileAccess.Write, FileShare.None);
			streamWriter = new StreamWriter(fileStream);
		}

		~PresentationStreamWriter()
		{
			if (File.Exists(fileBody))
			{
				File.Delete(fileBody);
			}
		}

		#endregion Ctor/Dtor

		#region Public Methods

		/// <summary>
		/// Create a new file with header.  Note: the caller is responsible for cleaning up this file.
		/// </summary>
		/// <param name="startTime"></param>
		/// <param name="baseUrl"></param>
		/// <param name="extent"></param>
		/// <returns>New file path, or null</returns>
		public String WriteHeaderAndCopy(String startTime, String baseUrl, String extent, string lectureTitle)
		{
			if (streamWriter != null)
			{
				return null;
			}

			DateTime start = DateTime.Parse(startTime);
			String fileName = Utility.GetTempFilePath("xml");

			FileStream fileStream = new FileStream(fileName,
				FileMode.Append, FileAccess.Write, FileShare.None);
         
			String checkedBaseUrl;
			if ((baseUrl==null) || (baseUrl.Trim() == ""))
			{
				log.Append("Job does not include a base url for slide images.  Slides will not be available to " +
					"webviewer in the streaming scenario until the URL base has been added to the presentation data (xml) file.");
				log.ErrorLevel = 3;
				checkedBaseUrl = "Enter URL base here";
			}
			else
			{
				if (baseUrl.Trim().EndsWith("/"))
				{
					checkedBaseUrl = baseUrl.Trim();
				}
				else
				{
					checkedBaseUrl = baseUrl.Trim() + "/";
				}
			}

			StreamWriter streamWriter2 = new StreamWriter(fileStream);
			streamWriter2.WriteLine("<?xml version=\"1.0\"?> ");
			streamWriter2.WriteLine("<!-- CXP ArchiveTranscoder: " + DateTime.Now.ToString() + " -->");
			streamWriter2.WriteLine("<WMBasicEdit > ");
			streamWriter2.WriteLine("<RemoveAllScripts /> ");
			streamWriter2.WriteLine("<ScriptOffset Start=\"" + start.ToString("M/d/yyyy HH:mm:ss.ff") + "\" /> ");
			streamWriter2.WriteLine("<Options NoAutoTOC=\"false\" PreferredWebViewerVersion=\"1.9.2.0\" /> ");
            streamWriter2.WriteLine("<Slides BaseURL=\"" + checkedBaseUrl + "\" Extent=\"" + extent + "\" />");
            streamWriter2.WriteLine("<Metadata Title=\"" + lectureTitle + "\" />");
            streamWriter2.Flush();

			//append fileBody to the end of fileName.  Return fileName.
			FileStream fsBody = File.OpenRead(fileBody);
			byte[] buf = new byte[fsBody.Length];
			fsBody.Read(buf,0,(int)fsBody.Length);
			fileStream.Write(buf,0,(int)fsBody.Length);
			fsBody.Close();

			streamWriter2.Flush();
			streamWriter2.Close();
			streamWriter2 = null;
			
			return fileName;
		}

		/// <summary>
		/// 
		/// </summary>
		public void Close()
		{
			if (streamWriter != null)
			{
				streamWriter.WriteLine( "</WMBasicEdit> ");
				streamWriter.Flush();
				streamWriter.Close();
				streamWriter = null;
			}
		}

		public void Flush()
		{
			if (streamWriter != null)
			{
				streamWriter.Flush();
			}
		}


		public void Write(string data)
		{
			if (streamWriter != null) 
			{
				lock (streamWriter)
					streamWriter.Write(data);
			}
		}

		#endregion Public Methods
	}
}
