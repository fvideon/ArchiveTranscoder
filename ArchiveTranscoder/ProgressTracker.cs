using System;
using System.Threading;
using System.Diagnostics;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Accept inputs and compose a status message.  Maintain a thread to display the message periodically.
	/// </summary>
	/// Message looks like:
	///		Segment 1 of 2: Reading Video 30% complete
	/// or:
	///		Writing Presentation Data: 10% complete
	///	"Writing Presentation Data" is a customMessage.  If it is set, it overrides the normal message.  
	///	segmentCount and currentSegment define the Segment data.  "Reading Video" is the avStatusMessage.  
	///	The percentage is calculated from currentValue and endValue.  If ShowPercentComplete is set to false
	///	that part will be omitted.
	public class ProgressTracker:IDisposable
	{
		#region Members

		private int segmentCount;
		private bool started;
		private Thread trackerThread;
		private bool endNow;
		private String avStatusMessage;
		private String customMessage;
		private int currentSegment;
		private int currentValue;
		private int endValue;
		private bool showPercentComplete;

		#endregion Members

		#region Ctor

		/// <summary>
		/// Create the ProgressTracker. segmentCount is the total number of segments.  In a message such as
		/// "Segment 1 of 2: Reading Video 30% complete", segmentCount is 2.
		/// </summary>
		/// <param name="segmentCount"></param>
		public ProgressTracker(int segmentCount)
		{
			this.segmentCount = segmentCount;
			started = false;
			avStatusMessage = "";
			customMessage = "";
			currentSegment = 0;
			currentValue = 0;
			endValue = 0;
			trackerThread = null;
			showPercentComplete = true;
		}

		#endregion Ctor

		#region IDisposable Members

		public void Dispose()
		{
			if (trackerThread != null)
			{
				endNow = true;
				if (!trackerThread.Join(1000))
				{
					trackerThread.Abort();
				}
				trackerThread = null;
			}
		}

		#endregion	
		
		#region Properties

		/// <summary>
		/// In a output string such as "Segment 1 of 2: Reading Video 30% complete", The substring "Reading Video"
		/// is the AVStatusMessage.  This will not show if CustomMessage is set.
		/// </summary>
		public String AVStatusMessage
		{
			set { avStatusMessage = value; }
		}

		/// <summary>
		/// Causes the output string to take the form such as "Writing Presentation Data: 10% complete".
		/// Here "Writing Presentation Data" is the Custom Message.  The CustomMessage overrides the
		/// AVStatusMessage.
		/// </summary>
		public String CustomMessage
		{
			set { customMessage = value; }
		}

		/// <summary>
		/// Only relevant if the CustomMessage is not set.  In a message such as this
		/// "Segment 1 of 2: Reading Video 30% complete", CurrentSegment is 1.
		/// </summary>
		public int CurrentSegment
		{
			set { currentSegment = value; }
		}

		/// <summary>
		/// Used to calculate Percent Complete.  Percent Complete = (100*CurrentValue)/EndValue.
		/// Percent Complete will not be displayed if ShowPercentComplete is set to false.
		/// </summary>
		public int CurrentValue
		{
			get { return currentValue; }
			set 
			{ 
				if (value > endValue)
				{
					currentValue = endValue;
				}
				else
				{
					currentValue = value;
				}
			}
		}

		/// <summary>
		/// Used to calculate Percent Complete.  Percent Complete = (100*CurrentValue)/EndValue.
		/// Percent Complete will not be displayed if ShowPercentComplete is set to false.
		/// </summary>
		public int EndValue
		{
			set 
            { 
                endValue = value; 
            }
            get { return endValue; }
		}

		/// <summary>
		/// Indicate whether to display or not display the Percent Complete.
		/// </summary>
		public bool ShowPercentComplete
		{
			set { showPercentComplete = value; }
		}

		#endregion Properties

		#region Public Methods

		/// <summary>
		/// Start the progress tracker thread to raise events periodically to update the message string.
		/// </summary>
		public void Start()
		{
			if (started)
			{
				Debug.WriteLine("Tracker Thread already started.");
				return;
			}
			endNow = false;
			trackerThread = new Thread(new ThreadStart(TrackerThread));
			trackerThread.Name = "Progress Tracker Thread";
			trackerThread.Start();
			started = true;
		}

		/// <summary>
		/// Stop the progress tracker thread.
		/// </summary>
		public void Stop()
		{
			if (trackerThread == null)
			{
				Debug.WriteLine("TrackerThread already killed.");
				return;
			}

			endNow = true;
			if (!trackerThread.Join(1000))
			{
				Debug.WriteLine("ProgressTracker thread aborting.");
				trackerThread.Abort();
			}

			if (OnShowProgress != null)
			{
				OnShowProgress("");
			}
			trackerThread = null;
			started = false;
		}

		#endregion Public Methods

		#region Private Methods

		/// <summary>
		/// Build the current message string.
		/// </summary>
		/// <returns></returns>
		private String CalculateMessage()
		{
			if (!started)
				return "Status: idle.";

			int percentage = 0;
			if ((endValue != 0) && (endValue >= currentValue))
			{
				percentage = (int)(((double)currentValue/(double)endValue) * 100.0);
			}

			String ret = "";
			if (customMessage != "")
			{
				ret = customMessage;
				if (showPercentComplete)
				{
					ret += ": " + percentage.ToString() + "% completed";
				} 
			}
			else
			{
				ret = "Segment " + currentSegment.ToString() + " of " + segmentCount.ToString() + ": " +
					avStatusMessage;
				if (showPercentComplete)
				{
					ret += ": " + percentage.ToString() + "% completed";
				} 
			}
			return ret;
		}

		/// <summary>
		/// Thread proc.  Raise event to show updates once per second.
		/// </summary>
		private void TrackerThread()
		{
			int repCount = 4;
			while (!endNow)
			{
				Thread.Sleep(200);
				repCount++;
				if (repCount>=5)
				{
					//Debug.WriteLine("Progress tracker reporting..");
					repCount = 0;
					if (OnShowProgress != null)
					{
						if (endNow)
							break;
						OnShowProgress(CalculateMessage());
					}
				}
			}
		}

		#endregion Private Methods

		#region Event

		/// <summary>
		/// Event to be raised when it is time to display a message.
		/// </summary>
		public event showProgressCallback OnShowProgress;
		public delegate void showProgressCallback(String message);

		#endregion Event
	}
}
