using System;
using MSR.LST.RTDocuments;
using System.Diagnostics;

namespace ArchiveTranscoder
{
	/// <summary>
	/// A few of the RTDocument utility functions from the MSR code recast here as static methods for convenience.
	/// </summary>
	public class RTDocumentUtility
	{
		public RTDocumentUtility()
		{}

		public static int PageIDToPageIndex(RTDocument rtDoc, Guid pageID)
		{
			foreach (TOCNode n in rtDoc.Organization.TableOfContents)
			{
				if (n.ResourceIdentifier == pageID)
				{
					return rtDoc.Organization.TableOfContents.IndexOf(n);
				}
			}

			return -1;
		}

		public static bool AddPageToRTDocument(RTDocument rtDocument, Page page)
		{
			TOCNode tn = WalkTOCNodes(rtDocument.Organization.TableOfContents, page);
			if (tn == null)
			{
				return false;
			}

            if (!rtDocument.Resources.Pages.ContainsKey(page.Identifier))
            {
                rtDocument.Resources.Pages.Add(page.Identifier, page);
                tn.Resource = page;
            }
			return true;
		}

		public static bool AddRTPageAddToRTDocument(RTDocument rtDocument, RTPageAdd rtpa)
		{
			rtDocument.Organization.TableOfContents.Add(rtpa.TOCNode);
			return AddPageToRTDocument(rtDocument,rtpa.Page);
		}

		public static bool AddPageIDToRTDocument(RTDocument rtDocument, Guid pageID)
		{
			TOCNode tocNode = new TOCNode();
			tocNode.ResourceIdentifier = pageID;
			if (rtDocument.Identifier==pageID)
			{
				tocNode.Identifier=pageID;
			}
			else
			{
				tocNode.Identifier = Guid.NewGuid();
			}
			Page p = new Page();
			p.Identifier = pageID;
			p.Image = null;
			tocNode.Resource = p;
			rtDocument.Organization.TableOfContents.Add(tocNode);
			return AddPageToRTDocument(rtDocument,p);
		}


		public static bool AddTocIDToRTDocument(RTDocument rtDocument, Guid tocID)
		{
			TOCNode tocNode = new TOCNode();
			tocNode.ResourceIdentifier = Guid.Empty;
			tocNode.Identifier = tocID;
			tocNode.Resource = null;
			rtDocument.Organization.TableOfContents.Add(tocNode);
			return true;
		}


		private static TOCNode WalkTOCNodes ( TOCList tocNodes, Page page )
		{
			foreach ( TOCNode tocNode in tocNodes )
			{
				if ( tocNode.ResourceIdentifier == page.Identifier )
				{
					return tocNode;
				}

				if ( tocNode.Children != null )
				{
					TOCNode tocNodeFromChildren = WalkTOCNodes ( tocNode.Children, page );

					if (tocNodeFromChildren != null)
					{
						return tocNodeFromChildren;
					}
				}
			}

			return null;
		}

	}
}
