using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RemoveDuplicatesNOW
{
    [ComVisible(true)]
    public class RemoveDuplicatesNOWRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public RemoveDuplicatesNOWRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("RemoveDuplicatesNOW.RemoveDuplicatesNOWRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnCalendarButton(Office.IRibbonControl control)
        {
            Outlook.Application outlookApplication = new Outlook.Application();
            NameSpace mapiNamespace = outlookApplication.GetNamespace("MAPI"); ;
            MAPIFolder CalendarFolder = mapiNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            Items outlookCalendarItems = CalendarFolder.Items;
            outlookCalendarItems.IncludeRecurrences = true;

            IList<AppointmentItem> items = new List<AppointmentItem>();
            IList<AppointmentItem> toRemove = new List<AppointmentItem>();

            foreach (AppointmentItem item in outlookCalendarItems)
            {
                if (items.Contains(item, new AppointmentItemComparer()))
                {
                    toRemove.Add(item);
                }
                else
                {
                    items.Add(item);
                }
            }

            foreach (var toRemoveItem in toRemove)
            {
                toRemoveItem.Delete();
            }

            MessageBox.Show(toRemove.Count + " Appointments removed.");
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }

    public class AppointmentItemComparer : IEqualityComparer<AppointmentItem>
    {

        public bool Equals(AppointmentItem x, AppointmentItem y)
        {
            if (x.Subject == y.Subject && x.Start == y.Start && x.End == y.End && x.Body == y.Body)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public int GetHashCode(AppointmentItem obj)
        {
            return obj.GetHashCode();
        }
    }
}
