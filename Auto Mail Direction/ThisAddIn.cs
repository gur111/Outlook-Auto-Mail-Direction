using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace Auto_Mail_Direction
{
    public partial class ThisAddIn
    {
        /*
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }*/

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.NewMailEx += new Outlook.ApplicationEvents_11_NewMailExEventHandler(olApp_NewMail);
        }

        private void olApp_NewMail(String entryIDCollection)
        {
            Outlook.NameSpace outlookNS = this.Application.GetNamespace("MAPI");
            Outlook.MAPIFolder mFolder = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.MailItem mail;

            try
            {
                mail = (Outlook.MailItem)outlookNS.GetItemFromID(entryIDCollection, Type.Missing);
                mail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
//                    mail.HTMLBody = "<html><body dir='auto'>" + mail.HTMLBody + "</body></html>";
                WebBrowser browser = new WebBrowser();
                browser.ScriptErrorsSuppressed = true;
                browser.DocumentText = mail.HTMLBody;
                browser.Document.OpenNew(true);
                browser.Document.Write(mail.HTMLBody);
                browser.Refresh();
                HtmlDocument doc =  browser.Document;
                SetDirections(doc);
                mail.HTMLBody = doc.Body.InnerHtml;
                Console.WriteLine(mail.HTMLBody);
                mail.Save();
            }
            catch
            { }
        }

        private void SetDirections(HtmlDocument doc) {
            HtmlElementCollection collection = doc.Body.All;
            foreach (HtmlElement elem in collection) {
                if (elem.InnerText != null && elem.InnerText.Length > 2) {
                    char[] text = elem.InnerText.ToArray();
                    int engCount = 0;
                    int hebCount = 0;
                    for (int i = 0; i < text.Length; i++) {
                        if ((text[i] >= 'a' && text[i] <= 'z') || (text[i] >= 'A' && text[i] <= 'Z')) {
                            engCount++;
                        } else if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(text[i]) == System.Globalization.UnicodeCategory.OtherLetter) {
                            hebCount++;
                        }
                    }
                    if (engCount + hebCount > 10)
                    {
                        if (1.1 * engCount >= hebCount && GetDirection(elem) != "ltr")
                        {
                            SetDirection(doc, elem, "ltr");
                        }
                        else if (1.1 * engCount < hebCount && GetDirection(elem) != "rtl")
                        {
                            SetDirection(doc, elem, "rtl");
                        }
                    }
                }
            }

        }

        private void SetDirection(HtmlDocument doc, HtmlElement elem, string dir) {
            if (IsDirectionable(elem.TagName.ToLower()))
            {
                elem.SetAttribute("dir", dir);
            }
            else {
                HtmlElement wrapper = doc.CreateElement("DIV");
                wrapper.SetAttribute("dir", dir);
                wrapper.InnerHtml = elem.OuterHtml;
                elem.OuterHtml = wrapper.OuterHtml;
            }
        }

        private bool IsDirectionable(string tagName) {
            return "html body div p span header table td tr a".Contains(tagName);
        }

        private string GetDirection(HtmlElement elem) {
            while (elem.Parent != null) {
                elem = elem.Parent;
                if (elem.GetAttribute("dir") == "rtl" || elem.GetAttribute("dir") == "ltr") {
                    return elem.GetAttribute("dir");
                }
            }
            return elem.GetAttribute("dir");
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
