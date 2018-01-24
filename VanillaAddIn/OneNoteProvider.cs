using Microsoft.Office.Interop.OneNote;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace JTools
{
    public class OneNoteNotebook
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string NickName { get; set; }
        public string Path { get; set; }
        public Color Color { get; set; }
    }

    public class OneNoteSection
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string Path { get; set; }
        public bool Encrypted { get; set; }
        public Color Color { get; set; }
    }

    public class OneNoteSectionGroup
    {
        public string Name { get; set; }
        public string ID { get; set; }
        public DateTime LastModified { get; set; }
        public string Path { get; set; }
    }

    public interface IOneNotePage
    {
        string ID { get; set; }
        string Name { get; set; }
        DateTime DateTime { get; set; }
        DateTime LastModified { get; set; }
    }

    public class OneNotePage
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public DateTime DateTime { get; set; }
        public DateTime LastModified { get; set; }
    }

    public interface IOneNoteExtNotebook { }

    public class OneNoteExtNotebook : OneNoteNotebook
    {
        public IEnumerable<OneNoteExtSection> Sections { get; set; }
        public IEnumerable<OneNoteSectionGroup> SectionGroups { get; set; }
    }

    public struct CurrentPage
    {
        public string ID;
        public string Name;
        public string SectionName;
        public string SectionID;
        public string SectionGroupName;
        public string NotebookName;
    }

    public interface IOneNoteExtSection { }

    public class OneNoteExtSection : OneNoteSection
    {
        public IEnumerable<OneNotePage> Pages { get; set; }
    }

    public interface IOneNoteExtPage : IOneNotePage
    {
        OneNoteSection Section { get; set; }
        OneNoteNotebook Notebook { get; set; }
    }

    public class OneNoteExtPage : OneNotePage
    {
        public OneNoteSection Section { get; set; }
        public OneNoteNotebook Notebook { get; set; }
    }

    /// <summary>
    /// OneNote Provider (LINQ to OneNote) 
    /// </summary>
    /// <remarks>
    /// <list type="bullet">
    /// <item>Author: Stefan Cruysberghs</item>
    /// <item>Website: http://www.scip.be</item>
    /// <item>Article: Querying Outlook and OneNote with LINQ : http://www.scip.be/index.php?Page=ArticlesNET05</item>
    /// <item>Source code: http://www.scip.be/index.php?Page=ComponentsNETOfficeItems</item>
    /// </list>
    /// </remarks>  
    public class OneNoteProvider
    {

        private readonly Application oneNote;
        private XNamespace ns = null;

        public XNamespace GetNamespace()
        {
            string xml;
            OneNote.GetHierarchy(null, HierarchyScope.hsNotebooks, out xml);

            var doc = XDocument.Parse(xml);
            ns = doc.Root.Name.Namespace;

            return ns;
        }

        /// <summary>
        /// Instance of OneNote Application object
        /// </summary>
        /// <remarks>
        /// <list type="bullet">
        /// <item>http://msdn.microsoft.com/en-us/library/ms788684.aspx</item>
        /// <item>http://msdn.microsoft.com/en-us/library/aa286798.aspx</item>
        /// </list>
        /// </remarks>
        public Application OneNote
        {
            get { return oneNote; }
        }

        /// <summary>
        /// Hierarchy of Notebooks with Sections and Pages
        /// </summary>
        /// <example><code>
        /// // All sections which are encrypted
        /// var queryEncryptedSections = from nb in oneNoteProvider.NotebookItems
        ///                              from s in nb.Sections
        ///                              where s.Encrypted == true
        ///                              select new { NotebookName = nb.Name, SectionName = s.Name };
        ///
        /// // All notebooks and the number of sections they have
        /// var queryNotebooks = (from nb in oneNoteProvider.NotebookItems
        ///                       select new { Notebook = nb.Name, SectionCount = nb.Sections.Count() })
        ///                      .OrderByDescending(n =&gt; n.SectionCount);
        /// </code></example>     
        public IEnumerable<OneNoteExtNotebook> NotebookItems
        {
            get
            {
                XElement oneNoteHierarchy = XElement.Parse(oneNoteXMLHierarchy);
                XNamespace one = oneNoteHierarchy.GetNamespaceOfPrefix("one");

                // Transform XML into object hierarchy
                IEnumerable<OneNoteExtNotebook> oneNoteNotebookItems = from n in oneNoteHierarchy.Elements(one + "Notebook")
                                                                       where n.HasAttributes
                                                                       select new OneNoteExtNotebook()
                                                                       {
                                                                           ID = n.Attribute("ID").Value,
                                                                           Name = n.Attribute("name").Value,
                                                                           NickName = n.Attribute("nickname").Value,
                                                                           Path = n.Attribute("path").Value,
                                                                           Color = ColorTranslator.FromHtml(n.Attribute("color").Value),
                                                                           Sections = n.Elements(one + "Section").Select(s => new OneNoteExtSection()
                                                                           {
                                                                               ID = s.Attribute("ID").Value,
                                                                               Name = s.Attribute("name").Value,
                                                                               Path = s.Attribute("path").Value,
                                                                               Color = ColorTranslator.FromHtml(s.Attribute("color").Value),
                                                                               Encrypted = ((s.Attribute("encrypted") != null) && (s.Attribute("encrypted").Value == "true")),
                                                                               Pages = s.Elements().Select(p => new OneNotePage()
                                                                               {
                                                                                   ID = p.Attribute("ID").Value,
                                                                                   Name = p.Attribute("name").Value,
                                                                                   DateTime = XmlConvert.ToDateTime(p.Attribute("dateTime").ToString(), XmlDateTimeSerializationMode.Utc),
                                                                                   LastModified = XmlConvert.ToDateTime(p.Attribute("lastModifiedTime").ToString(), XmlDateTimeSerializationMode.Utc)
                                                                               }).OfType<OneNotePage>()
                                                                           }).OfType<OneNoteExtSection>(),
                                                                           SectionGroups = n.Elements(one + "SectionGroup").Select(g => new OneNoteSectionGroup()
                                                                           {
                                                                               ID = g.Attribute("ID").Value,
                                                                               Name = g.Attribute("name").Value,
                                                                               Path = g.Attribute("path").Value,
                                                                               LastModified = XmlConvert.ToDateTime(g.Attribute("lastModifiedTime").ToString(), XmlDateTimeSerializationMode.Utc)
                                                                           }).OfType<OneNoteSectionGroup>()
                                                                       };

                return oneNoteNotebookItems;
            }
        }

        /// <summary>
        /// Collection of Pages
        /// </summary>
        /// <example><code>
        /// // Query all pages which have been modified last month
        /// var queryPages = from page in oneNoteProvider.PageItems
        ///                  where page.LastModified &gt; DateTime.Now.AddMonths(-1)
        ///                  orderby page.LastModified descending
        ///                  select page;
        /// 
        /// // Show XML content of pages which have been changed yesterday
        /// foreach (var item in oneNoteProvider.PageItems.Where(p =&gt; p.LastModified &gt; DateTime.Now.AddDays(-1)))
        /// {
        ///   string pageXMLContent = "";
        ///   Console.WriteLine("{0} {1} {2} {3} {4}", item.LastModified, item.Notebook.Name, item.Section.Name, item.Name, item.DateTime);
        ///   oneNoteProvider.OneNote.GetPageContent(item.ID, out pageXMLContent, Microsoft.Office.Interop.OneNote.PageInfo.piBasic);
        ///   Console.WriteLine("{0}", pageXMLContent);
        /// }
        /// </code></example> 
        public IEnumerable<OneNoteExtPage> PageItems
        {
            get
            {
                XElement oneNoteHierarchy = XElement.Parse(oneNoteXMLHierarchy);
                XNamespace one = oneNoteHierarchy.GetNamespaceOfPrefix("one");

                // Transform XML into object collection
                IEnumerable<OneNoteExtPage> oneNotePageItems = from o in oneNoteHierarchy.Elements(one + "Notebook").Elements(one + "Section").Elements(one + "Page")
                                                               where o.HasAttributes
                                                               select new OneNoteExtPage()
                                                               {
                                                                   ID = o.Attribute("ID").Value,
                                                                   Name = o.Attribute("name").Value,
                                                                   DateTime = XmlConvert.ToDateTime(o.Attribute("dateTime").Value, XmlDateTimeSerializationMode.Utc),
                                                                   LastModified = XmlConvert.ToDateTime(o.Attribute("lastModifiedTime").Value, XmlDateTimeSerializationMode.Utc),
                                                                   Section = new OneNoteSection()
                                                                   {
                                                                       ID = o.Parent.Attribute("ID").Value,
                                                                       Name = o.Parent.Attribute("name").Value,
                                                                       Path = o.Parent.Attribute("path").Value,
                                                                       Color = ColorTranslator.FromHtml(o.Parent.Attribute("color").Value),
                                                                       Encrypted = ((o.Parent.Attribute("encrypted") != null) && (o.Parent.Attribute("encrypted").Value == "true"))
                                                                   },
                                                                   Notebook = new OneNoteNotebook()
                                                                   {
                                                                       ID = o.Parent.Parent.Attribute("ID").Value,
                                                                       Name = o.Parent.Parent.Attribute("name").Value,
                                                                       NickName = o.Parent.Parent.Attribute("nickname").Value,
                                                                       Path = o.Parent.Parent.Attribute("path").Value,
                                                                       Color = ColorTranslator.FromHtml(o.Parent.Parent.Attribute("color").Value)
                                                                   }
                                                               };

                return oneNotePageItems;
            }
        }

        public string GetObjectId(string parentId, HierarchyScope scope, string objectName)
        {
            string xml;
            OneNote.GetHierarchy(parentId, scope, out xml);

            var doc = XDocument.Parse(xml);
            var nodeName = "";

            switch (scope)
            {
                case (HierarchyScope.hsNotebooks): nodeName = "Notebook"; break;
                case (HierarchyScope.hsPages): nodeName = "Page"; break;
                case (HierarchyScope.hsSections): nodeName = "Section"; break;
                default:
                    return null;
            }

            var node = doc.Descendants(ns + nodeName).Where(n => n.Attribute("name").Value == objectName).FirstOrDefault();

            return node.Attribute("ID").Value;
        }

        public string GetPageMLOTagDef(string pageContentXML)
        {
            XElement oneNoteHierarchy = XElement.Parse(pageContentXML);
            XNamespace one = oneNoteHierarchy.GetNamespaceOfPrefix("one");

            var mloTagDefIndex = from t in oneNoteHierarchy.Elements(one + "TagDef")
                                 where t.Attribute("name").Value == "MLO"
                                 select t.Attribute("index");

            return mloTagDefIndex.First().Value;
        }

        public IEnumerable<XElement> GetMLOTaggedElements(string pageContentXML, string MLOTagIndex)
        {
            XElement oneNoteHierarchy = XElement.Parse(pageContentXML);
            XNamespace one = oneNoteHierarchy.GetNamespaceOfPrefix("one");

            var elements = from oe in oneNoteHierarchy.Descendants(one + "Tag")              
                           where oe.Attribute("index").Value == MLOTagIndex
                           where oe.Attribute("completed").Value == "false"
                           select oe.Parent;

            return elements;
        }

        public void SetTagValue(string pageID, string OEObjectID, string tagDefID)
        {
            string oldPageContent = GetPageContent(pageID);
            XElement oneNoteHierarchy = XElement.Parse(oldPageContent);
            XNamespace one = oneNoteHierarchy.GetNamespaceOfPrefix("one");

            var element = from oe in oneNoteHierarchy.Descendants(one + "OE")
                          where oe.Attribute("objectID").Value == OEObjectID
                          select oe;

            var tag = element.Elements(one + "Tag").Where(t => t.Attribute("index").Value == tagDefID).First();

            tag.Attribute("completed").Value = "true";

            string newPageContent = tag.Ancestors(one + "Page").First().ToString();

            //string xml = "<?xml version=\"1.0\"><one:Page ID=\"" + pageID + "\">" + tag.Ancestors(one + "OE").First().ToString() + "</one:Page>";

            SetPageContent(newPageContent);
            //SetPageContent(xml);
        }

        public string GetScript(XElement element)
        {
            XNamespace one = element.GetNamespaceOfPrefix("one");

            var inkwords = from ink in element.Elements(one + "InkWord")
                           select ink.Attribute("recognizedText");

            StringBuilder recognisedText = new StringBuilder();

            foreach (var recognisedWord in inkwords)
            {
                if (recognisedWord != null)
                {
                    recognisedText.Append(recognisedWord.Value);
                    recognisedText.Append(" ");
                }
            }

            return recognisedText.ToString();
        }

        public string GetText(XElement element)
        {
            XNamespace one = element.GetNamespaceOfPrefix("one");

            var text = element.Element(one + "T");

            if (text != null)
                return text.Value;
            else
                return string.Empty;
        }

        public CurrentPage CurrentlyViewedPage
        {
            get
            {
                oneNote.GetHierarchy(null, HierarchyScope.hsPages, out oneNoteXMLHierarchy);

                var currentPage = new CurrentPage();

                XElement oneNoteHierarchy = XElement.Parse(oneNoteXMLHierarchy);
                XNamespace one = oneNoteHierarchy.GetNamespaceOfPrefix("one");

                var activePage = from p in oneNoteHierarchy.Descendants(one + "Page")
                                  where p.Attribute("isCurrentlyViewed") != null
                                  select p;

                currentPage.ID = activePage.First().Attribute("ID").Value;
                currentPage.Name = activePage.First().Attribute("name").Value;

                currentPage.SectionName = activePage.Ancestors(one + "Section").First().Attribute("name").Value;
                currentPage.SectionID = activePage.Ancestors(one + "Section").First().Attribute("ID").Value;

                if (activePage.Ancestors(one + "SectionGroup").Elements().Count() > 0)
                    currentPage.SectionGroupName = activePage.Ancestors(one + "SectionGroup").First().Attribute("name").Value;
                
                currentPage.NotebookName = activePage.Ancestors(one + "Notebook").First().Attribute("name").Value;

                return currentPage;
            }
        }


        public string GetPageContent(string pageID)
        {
            string xml;
            oneNote.GetPageContent(pageID, out xml, PageInfo.piBasic);

            return xml;
        }

        public string GetLinkToTag(string PageId, string TagId)
        {
            string link;

            oneNote.GetHyperlinkToObject(PageId, TagId, out link);

            return link;
        }

        public void SetPageContent(string newContent)
        {            
            oneNote.UpdatePageContent(newContent, DateTime.MinValue);
        }

        private string oneNoteXMLHierarchy = "";

        /// <summary>
        /// Constructor. Create instance of Microsoft.Office.Interop.OneNote.Application and get XML hierarchy.
        /// </summary>
        public OneNoteProvider()
        {
            oneNote = new Microsoft.Office.Interop.OneNote.Application();
            // Get OneNote hierarchy as XML document
            oneNote.GetHierarchy(null, HierarchyScope.hsPages, out oneNoteXMLHierarchy);
        }

        public void SetPageTitle(string pageID, string newTitle)
        {
            var pageContent = GetPageContent(pageID);

            XElement oneNoteHierarchy = XElement.Parse(pageContent);
            XNamespace one = oneNoteHierarchy.GetNamespaceOfPrefix("one");

            var title = oneNoteHierarchy.Elements(one + "Title").Descendants(one + "T").First();
            title.Value = newTitle;

            SetPageContent(oneNoteHierarchy.ToString());
        }

    }
}