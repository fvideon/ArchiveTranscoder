using System;
using System.Configuration;
using System.Reflection;
using System.Globalization;

namespace ArchiveTranscoder
{
	/// <summary>
	/// A place for application constants
	/// </summary>
	public class Constants
	{

        static Constants() {
            Type myType = typeof(Constants);
            FieldInfo[] fields = myType.GetFields(BindingFlags.Public | BindingFlags.Static);
            foreach (FieldInfo field in fields) {
                if (field.IsLiteral) // Is Constant?
                    continue; // We can't set it

                string fullName = field.Name;
                string configOverride = System.Configuration.ConfigurationManager.AppSettings[fullName];
                if (configOverride != null && configOverride != String.Empty) {
                    Type newType = field.FieldType;
                    object newValue = Convert.ChangeType(configOverride, newType, CultureInfo.InvariantCulture);
                    field.SetValue(null, newValue);
                }
            }
        }

		//public static String ConnectionStringTemplate = "Persist Security Info=False;Integrated Security=SSPI;database={1};server={0}";
        public static String SQLConnectionString = "Persist Security Info=False;Integrated Security=SSPI;database={1};server={0}";

		public static readonly int      TicksPerMs = 10000;                                     // handy constant
		public static readonly int      TicksPerSec = 10000000;                                 // handy constant
		public static readonly int      PlaybackBufferInterval = 10 * 1000 * TicksPerMs;        // temporal length of buffer, in ms
		public static readonly String	AppRegKey = "Software\\UW CSE\\ArchiveTranscoder";
		
        public static String TempDirectory = System.IO.Path.GetTempPath();

        /// <summary>
        /// If true, scale HD images down by reducing dimensions by half.  If false keep the source image size.
        /// </summary>
        public static bool ScaleFramegrabs = false;

        /// <summary>
        /// If true, the source is pillarboxed widescreen format from which we should attempt to restore a 
        /// 4x3 frame dimension.
        /// </summary>
        public static bool RemovePillarboxing = false;

		//Since we are storing DateTime as a string, we need enough precision to avoid roundoff errors
        //public static readonly string dtformat = "M/d/yyyy H:mm:ss.fff";
        //public static readonly string shortdtformat = "M/d/yyyy H:mm:ss.f";
        public static readonly string timeformat = "H:mm:ss.fff";
        public static readonly string shorttimeformat = "H:mm:ss.f";
		
		//When CP is used as a capability, it puts this identifier in the CP-specific RTFrames
		public readonly static Guid CLASSROOM_PRESENTER_MESSAGE = new Guid("{7D1D6F5F-D1D2-4b89-A510-861E97678CB5}");

        //CP uses this tag to index the ink extended properties
        public readonly static Guid CPInkExtendedPropertyTag = new Guid ("{179222D6-BCC1-4570-8D8F-7E8834C1DD2A}"); 

		//These guids identify the contents of CP messages in RTFrames:
		public readonly static Guid RTDocEraseAllGuid = new Guid("{E3472EA5-29BC-A4FE-25F3-B135F03C134D}");
		public readonly static Guid PageUpdateIdentifier = new Guid("{A3DCB171-6A51-405e-A977-CD3FC7AAA871}");
		public readonly static Guid StrokeDrawIdentifier = new Guid("{27B52DBB-A5DF-4cd7-9D8B-60ABD0D81BBC}");
		public readonly static Guid StrokeDeleteIdentifier = new Guid("{172E6122-9333-4ec0-9C23-E351038EBBD0}");

	}

    //Strings that are valid in the job segment flags array
    public enum SegmentFlags
    { 
        SlidesReplaceVideo
    }
}
