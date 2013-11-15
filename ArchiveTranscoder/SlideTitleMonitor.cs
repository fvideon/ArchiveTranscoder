using System;
using System.Collections;
using System.Security;
using System.Text;
using PresenterNav;
using System.Text.RegularExpressions;
using MSR.LST.RTDocuments;

namespace ArchiveTranscoder
{
	/// <summary>
	/// Watch for new slide decks being advertised, and store the deck and slide titles.
	/// Return the slide info in a XML-formatted string
	/// </summary>
	public class SlideTitleMonitor
	{
		#region Members

		private Hashtable decksStored; //key is deck guid, value is the ArrayList containing slide title strings
		private Hashtable deckTitles;  //key is deck guid, value is the string containing the deck title

		#endregion Members

		#region Properties

		public Hashtable DeckTitles
		{
			get {return deckTitles;}
		}

		#endregion Properties

		#region Ctor

		public SlideTitleMonitor()
		{
			decksStored = new Hashtable();
			deckTitles = new Hashtable();
		}

		#endregion Ctor

		#region Public Methods

		/// <summary>
		/// If the CPDeckCollection contains info about a deck that has not yet been seen, store the deck title
		/// and slide titles.
		/// </summary>
		/// <param name="obj"></param>
		public void ReceivePresenterNav(CPDeckCollection dc)
		{
			foreach (CPDeckSummary ds in dc.SummaryCollection)
			{
				if ((ds.DeckType == "Presentation") &&
					(ds.DocRef.GUID != Guid.Empty) &&
					(ds.SlideTitles.Count > 0))
				{
					if (!decksStored.ContainsKey(ds.DocRef.GUID))
						decksStored.Add(ds.DocRef.GUID,ds.SlideTitles);
					if (!deckTitles.ContainsKey(ds.DocRef.GUID))
						deckTitles.Add(ds.DocRef.GUID,ds.FileName);
				}
			}
		}

		public void ReceiveRTPageAdd(RTPageAdd rtpa, Guid docId)
		{
			if (!decksStored.ContainsKey(docId))
			{
				ArrayList titles = new ArrayList();
				titles.Add(rtpa.TOCNode.Title);
				decksStored.Add(docId,titles);
			}
			else
			{
				((ArrayList)decksStored[docId]).Add(rtpa.TOCNode.Title);
			}
			
			if (!deckTitles.ContainsKey(docId))
			{
				deckTitles.Add(docId,"RTDocument Deck");
			}

		}

		public void ReceiveRTDocument(RTDocument rtd)
		{
			if (!decksStored.ContainsKey(rtd.Identifier))
			{
				ArrayList titles = new ArrayList();
				foreach (TOCNode n in rtd.Organization.TableOfContents)
				{
					titles.Add(n.Title);
				}

				decksStored.Add(rtd.Identifier,titles);
			}
			if (!deckTitles.ContainsKey(rtd.Identifier))
			{
				deckTitles.Add(rtd.Identifier,"RTDocument Deck");
			}
		}

		/// <summary>
		/// Return a string containing XML formatted slide info including deck Guid, slide indices and slide titles.
		/// </summary>
		/// <returns></returns>
		public String Print()
		{
			StringBuilder sb = new StringBuilder();
			ArrayList titlesList;
            foreach (Guid deckGuid in deckTitles.Keys) {
                sb.Append("<DeckTitle DeckGuid=\"" + deckGuid.ToString() + "\" Text=\"" + FixTitle((string)deckTitles[deckGuid]) + "\"/>\r\n");
            }
			foreach (Guid deckGuid in decksStored.Keys)
			{
				titlesList = (ArrayList)decksStored[deckGuid];
				for (int i=0;i<titlesList.Count;i++)
				{
					/// Log with zero-based index.
					sb.Append("<Title DeckGuid=\""+deckGuid.ToString() + "\" Index=\"" + i.ToString() +
						"\" Text=\"" + FixTitle((String)titlesList[i]) + "\"/>\r\n");
				}
			}
			return sb.ToString();
		}

		#endregion Public Methods

		#region Private Methods

		/// <summary>
		/// Make transformations on the title we're given to make a useful and XML compatible title.
		/// </summary>
		/// <param name="rawTitle"></param>
		/// <returns></returns>
		private String FixTitle(String rawTitle)
		{
			// The slide titles may be formed like this: "  1. Title".  Convert to "Title".
			String fixedTitle = Regex.Replace((rawTitle),@"^\s*[0-9]*\.\s","");
			// Trim any whitespace characters from front and end
			fixedTitle = fixedTitle.Trim();
			// Replace invalid XML characters with valid equivalents.
			fixedTitle = SecurityElement.Escape(fixedTitle);
			return fixedTitle;
		}

		#endregion Private Methods
	}
}
