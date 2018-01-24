/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 */

 using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Extensibility;
//using JTools;
using JTools.Utilities;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.OneNote;
using Microsoft.Win32;
using System.Linq;
using MyApplication.VanillaAddIn.Properties;
//using MyApplication.MLOAddIn.Utilities;
//using Application = Microsoft.Office.Interop.OneNote.Application;  // Conflicts with System.Windows.Forms

#pragma warning disable CS3003 // Type is not CLS-compliant

namespace JTools.MLO
{
	[ComVisible(true)]
	[Guid("1B1A0A1B-2FF8-4364-ABBF-6870EB7C13D6"), ProgId("OneNote.MLOAddIn")]
    
    public class AddIn : IDTExtensibility2, IRibbonExtensibility
	{
		//protected Application OneNoteApplication
		//{ get; set; }

		//private MainForm mainForm;

        OneNoteProvider oneNoteProvider;
        ArrayList changes = new ArrayList();

        string config_exe = string.Empty;
        string config_datafile = string.Empty;
        string config_rootguid = string.Empty;
        string config_logfile = string.Empty;
        bool config_isDebug = false;
        
        public AddIn()
		{
		}

		/// <summary>
		/// Returns the XML in Ribbon.xml so OneNote knows how to render our ribbon
		/// </summary>
		/// <param name="RibbonID"></param>
		/// <returns></returns>
		public string GetCustomUI(string RibbonID)
		{
			return MyApplication.VanillaAddIn.Properties.Resources.ribbon;
		}

		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>
		/// Cleanup
		/// </summary>
		/// <param name="custom"></param>
		public void OnBeginShutdown(ref Array custom)
		{
			//this.mainForm?.Invoke(new Action(() =>
			//{
			//	// close the form on the forms thread
			//	this.mainForm?.Close();
			//	this.mainForm = null;
			//}));
            oneNoteProvider = null;
        }

		/// <summary>
		/// Called upon startup.
		/// Keeps a reference to the current OneNote application object.
		/// </summary>
		/// <param name="application"></param>
		/// <param name="connectMode"></param>
		/// <param name="addInInst"></param>
		/// <param name="custom"></param>
		public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
		{
			//SetOneNoteApplication((Application)Application);
            try
            {
                //MessageBox.Show("Going in!");
                oneNoteProvider = new OneNoteProvider();
                //MessageBox.Show("JTools OnConnection successful", "JTools AddIn", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show(string.Format("Error in JTools onConnection: {0}", e.Message), "JTools AddIn", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogAction(string.Format("Error in JTools onConnection: {0}", e.Message));
            }
        }

		//public void SetOneNoteApplication(Application application)
		//{
		//	OneNoteApplication = application;
		//}

		/// <summary>
		/// Cleanup
		/// </summary>
		/// <param name="RemoveMode"></param>
		/// <param name="custom"></param>
		[SuppressMessage("Microsoft.Reliability", "CA2001:AvoidCallingProblematicMethods", MessageId = "System.GC.Collect")]
		public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
		{
			//OneNoteApplication = null;
            oneNoteProvider = null;
            GC.Collect();
			GC.WaitForPendingFinalizers();
		}

		public void OnStartupComplete(ref Array custom)
		{
            //MessageBox.Show("JTools startup completed", "JTools AddIn", MessageBoxButtons.OK, MessageBoxIcon.Information);
            LogAction(string.Format("JTools startup completed"));
        }

        //public async Task MLOAddInButtonClicked(IRibbonControl control)
        //{
        //	MessageBox.Show("MLOAddIn button pushed! Now we'll load up the full XML hierarchy as well as the current page XML. This may take some time.");
        //	ShowForm();
        //	return;
        //}

        //private void ShowForm()
        //{
        //	this.mainForm = new MainForm(this.OneNoteApplication);
        //          System.Windows.Forms.Application.Run(this.mainForm);
        //}

        public IStream GetImage(string imageName)
        {
            MemoryStream mem = new MemoryStream();

            switch (imageName)
            {
                case "MLO.png":
                    Resources.MLO.Save(mem, ImageFormat.Png);
                    break;

                case "ReTitle.png":
                    Resources.ReTitle.Save(mem, ImageFormat.Png);
                    break;

                case "Outline.png":
                    Resources.ReTitle.Save(mem, ImageFormat.Png);
                    break;

                case "setup_icon.png":
                    Resources.setup_icon.Save(mem, ImageFormat.Png);
                    break;
            }

            return new CCOMStreamWrapper(mem);
        }

        public async void LogAction(string action)
        {
            if (config_isDebug && File.Exists(config_logfile))
            {
                await WriteTextAsync(config_logfile, action);
            }
        }

        private async Task WriteTextAsync(string filePath, string text)
        {
            byte[] encodedText = System.Text.Encoding.Unicode.GetBytes(text);

            using (FileStream sourceStream = new FileStream(filePath,
                FileMode.Append, FileAccess.Write, FileShare.None,
                bufferSize: 4096, useAsync: true))
            {
                await sourceStream.WriteAsync(encodedText, 0, encodedText.Length);
            };
        }

        public void SetTitleTS(IRibbonControl control)
        {
            var currentPage = oneNoteProvider.CurrentlyViewedPage;

            oneNoteProvider.SetPageTitle(currentPage.ID, string.Format("{0} : {1}", currentPage.SectionName, DateTime.Now.ToShortDateString()));
        }

        public string GetMLOLinkToTag(string PageId, string TagId)
        {
            string link;

            link = oneNoteProvider.GetLinkToTag(PageId, TagId);

            return string.Format("file:{0}", link);
        }

        private void ReadSettings()
        {
            try
            {
                RegistryKey rk = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Ratsey\OneNoteMLO");

                config_exe = rk.GetValue("MLOExecutable", "").ToString();
                config_datafile = rk.GetValue("DataFile", "").ToString();
                config_rootguid = rk.GetValue("RootGUID", "").ToString();
                config_isDebug = rk.GetValue("IsDebug", "").ToString() == "True";
                config_logfile = rk.GetValue("LogFile", "").ToString();

                //MessageBox.Show("JTools settings read from registry", "JTools AddIn", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LogAction(string.Format("\n\nRead settings from registry:\n\tconfig_exe = {0}\n\tconfig_datafile = {1}\n\tconfig_rootguid = {2}", config_exe, config_datafile, config_rootguid));
            }
            catch (Exception e)
            {
                MessageBox.Show(string.Format("Error reading JTools registry settings: {0}", e.Message), "JTools AddIn", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogAction(string.Format("Error reading JTools registry settings: {0}", e.Message));
            }
        }

        public void SendCommandLine(string MLOPath, string DataFile, string RootTask, string NewTask, string NoteText, string PageTitle, string SectionName, string NotebookName)
        {
            string parsedNewTask = NewTask.Replace(@"""", @"""""");

            Process p = new Process();
            p.StartInfo.FileName = MLOPath;
            p.StartInfo.Arguments = string.Format(@"{0} -task={1} -AddSubtask=""{2} [{6}>{5}>{4}]({3})"" -console", DataFile, RootTask, parsedNewTask, NoteText, PageTitle, SectionName, NotebookName);

            LogAction(string.Format("\n\nSending MLO CLI:\n\t{0} {1}", MLOPath, p.StartInfo.Arguments));

            p.Start();

            Thread.Sleep(250);
        }
        
        public void showSetup(IRibbonControl control)
        {
            try
            {
                //MessageBox.Show("Showing JTools AddIn setup", "JTools AddIn", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Form setup = new frmSetup();
                setup.ShowDialog();
            }
            catch (Exception e)
            {
                MessageBox.Show(string.Format("Error in showing setup window: {0}", e.Message), "JTools AddIn", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogAction(string.Format("Error in showing setup window: {0}", e.Message));
            }
        }

        public void CreateMLOTasks(IRibbonControl control)
        {
            if (IsSetup())
            {

                var currentPage = oneNoteProvider.CurrentlyViewedPage;
                var pageContent = oneNoteProvider.GetPageContent(currentPage.ID);
                var tagIndex = oneNoteProvider.GetPageMLOTagDef(pageContent);
                var elements = oneNoteProvider.GetMLOTaggedElements(pageContent, tagIndex);

                foreach (var element in elements)
                {

                    string taggedItem = oneNoteProvider.GetScript(element) + oneNoteProvider.GetText(element);
                    string linkToTag = "file:" + oneNoteProvider.GetLinkToTag(currentPage.ID, element.Attribute("objectID").Value);

                    oneNoteProvider.SetTagValue(currentPage.ID, element.Attribute("objectID").Value, tagIndex);

                    SendCommandLine(config_exe,
                        config_datafile,
                        config_rootguid,
                        taggedItem,
                        linkToTag,
                        currentPage.Name,
                        currentPage.SectionName,
                        currentPage.NotebookName);
                }
            }
            else
            {
                Form setup = new frmSetup();
                setup.ShowDialog();
            }
        }

        private bool IsSetup()
        {
            ReadSettings();

            if (config_exe.Length > 0 && config_datafile.Length > 0 && config_rootguid.Length > 0)
                return true;
            else
                return false;
        }

        // TODO: Extend to not destroy current page content!
        [STAThread]
        public void CreateOutline(IRibbonControl control)
        {
            Thread newThread = new Thread(new ThreadStart(CreateOutlineWork));
            newThread.SetApartmentState(ApartmentState.STA);
            newThread.Start();
        }

        public string CreatePage(string sectionId, string pageName, XElement insert)
        {
            // Create the new page
            string pageId;
            XNamespace ns = oneNoteProvider.GetNamespace();
            oneNoteProvider.OneNote.CreateNewPage(sectionId, out pageId, NewPageStyle.npsBlankPageWithTitle);

            // Get the title and set it to our page name
            string xml;
            oneNoteProvider.OneNote.GetPageContent(pageId, out xml, PageInfo.piAll);
            var doc = XDocument.Parse(xml);
            var title = doc.Descendants(ns + "T").First();
            title.Value = pageName;

            //Clipboard.SetText(insert.ToString());

            doc.Root.AddFirst(new XElement(ns + "QuickStyleDef"));
            var def = doc.Descendants(ns + "Page").Descendants(ns + "QuickStyleDef").First();
            def.Add(new XAttribute("index", "1"));
            def.Add(new XAttribute("name", "h3"));
            def.Add(new XAttribute("fontColor", "#5B9BD5"));
            def.Add(new XAttribute("highlightColor", "automatic"));
            def.Add(new XAttribute("font", "Calibri"));
            def.Add(new XAttribute("fontSize", "12.0"));
            def.Add(new XAttribute("spaceBefore", "0.5"));
            def.Add(new XAttribute("spaceAfter", "0.0"));

            doc.Root.AddFirst(new XElement(ns + "TagDef"));
            var tag = doc.Descendants(ns + "Page").Descendants(ns + "TagDef").First();
            tag.Add(new XAttribute("index", "0"));
            tag.Add(new XAttribute("type", "1"));
            tag.Add(new XAttribute("symbol", "3"));
            tag.Add(new XAttribute("fontColor", "automatic"));
            tag.Add(new XAttribute("highlightColor", "none"));
            tag.Add(new XAttribute("name", "To Do"));

            doc.Descendants(ns + "Page").Last().Add(insert);

            // Update the page
            oneNoteProvider.OneNote.UpdatePageContent(doc.ToString());

            return pageId;
        }

        // Generate a new page containing an outline of tasks obtained from MLO's 'Copy Tasks' function
        private void CreateOutlineWork()
        {
            try
            {
                XNamespace ns = oneNoteProvider.GetNamespace();
                var currentPage = oneNoteProvider.CurrentlyViewedPage;
                var pageContent = oneNoteProvider.GetPageContent(currentPage.ID);

                var root = new XElement(ns + "Outline");

                root.Add(new XElement(ns + "OEChildren"));

                var clipContent = Clipboard.GetText(TextDataFormat.Text).Trim();
                var lines = clipContent.Split('\n');
                int i = 0;
                var line = " ";

                while (line.Length > 0 && i < lines.Length)
                {
                    line = lines[i];

                    root.Descendants(ns + "OEChildren").First().Add(new XElement(ns + "OE"));
                    var oe = root.Descendants(ns + "OEChildren").Descendants(ns + "OE").Last();

                    // Tag
                    if (line.StartsWith(" "))
                    {
                        oe.Add(new XElement(ns + "Tag"));
                        var item = oe.Descendants(ns + "Tag").Last();
                        item.Add(new XAttribute("index", "0"));
                        item.Add(new XAttribute("completed", "false"));
                        item.Add(new XAttribute("disabled", "false"));
                    }
                    else
                        oe.Add(new XAttribute("quickStyleIndex", "1"));

                    oe.Add(new XElement(ns + "T"));
                    var title = oe.Descendants(ns + "T").First();
                    title.Value = $"{line}";

                    i++;
                }


                // Create a new page
                var page = CreatePage(currentPage.SectionID, "MLO Import", root);

            }
            catch (Exception e)
            {
                LogAction(string.Format("Error parsing clipboard content: {0}", e.Message));
                MessageBox.Show(string.Format("Error parsing clipboard content to OneNote outline: {0}", e.Message), "JTools AddIn", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
