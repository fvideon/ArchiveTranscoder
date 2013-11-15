using System;
using System.Text;
using System.IO;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Wrap a StringBuilder and add features to support writing a log file.
	/// </summary>
	public class LogMgr
	{
		#region Members

		private StringBuilder log;
		private String tempFilePath;
		private int errorLevel;

		#endregion Members

		#region Ctor/Dtor

		public LogMgr()
		{
			log = new StringBuilder();
			tempFilePath = null;
			errorLevel = 0;
		}

		~LogMgr()
		{
			if(tempFilePath!=null)
			{
				if (File.Exists(tempFilePath))
				{
					File.Delete(tempFilePath);
				}
				tempFilePath = null;
			}
		}

		#endregion Ctor/Dtor

		#region Properties

		/// <summary>
		/// Value from 0 - 10 which indicates the severity of an error or warning.
		/// This value can only ratchet up for the lifetime of the log.  Its value
		/// will tell us the worst severity (highest value) entry logged so we can
		/// determine how to inform the user. 
        /// {0-4}:Information;{5-6}:Warning;{7-10}:Error.
		/// </summary>
		public int ErrorLevel
		{
			get { return errorLevel; }
			set 
			{ 
				if ((value >=0) && (value <= 10))
				{
					if (value > errorLevel)
						errorLevel = value; 
				}
			}
		}

		#endregion Properties

		#region Public Methods

		/// <summary>
		/// Append the LogMgr to this one and transfer the ErrorLevel value as well.
		/// </summary>
		/// <param name="logMgr"></param>
		public void Append(LogMgr logMgr)
		{
			log.Append(logMgr.ToString());
			this.ErrorLevel = logMgr.ErrorLevel;
		}

		public void Append(String line)
		{
			log.Append(line);
		}

		public void WriteLine(String line)
		{
			log.Append(Utility.GetLocalizedDateTimeString(DateTime.Now, Constants.shorttimeformat) + ": " + line + "\r\n");
		}

		public String WriteTempFile()
		{
			if(tempFilePath!=null)
			{
				if (File.Exists(tempFilePath))
				{
					File.Delete(tempFilePath);
				}
			}
			tempFilePath = Utility.GetTempFilePath("txt");
			
			StreamWriter logSw = File.CreateText(tempFilePath);
			logSw.Write(log.ToString());
			logSw.Close();
			
			return tempFilePath;
		}
		
		public override String ToString()
		{
			return log.ToString();
		}

		#endregion Public Methods
	}
}
