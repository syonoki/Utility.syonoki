using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Outlook;

namespace Utility.syonoki.MSOffice {
    public static class OutlookSimpleApi {
        public static Application application { get; } = new Microsoft.Office.Interop.Outlook.Application();

        public static MAPIFolder getMailFolder(string olfolderName) {
            NameSpace ns = application.GetNamespace("mapi");
            return ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Folders[olfolderName];
        }

        public static bool hasMailReceived(string olFolderName, string containedSubject
            , out MailItem mail, DateTime date = new DateTime()) {
            if (date == new DateTime()) date = DateTime.Today;

            mail = findMailItemOn(olFolderName, date, containedSubject).FirstOrDefault();
            return mail != null;
        }

        public static IEnumerable<MailItem> findMailItemsBetween(string olFolderName, DateTime begin, DateTime end, string subject = null) {
            if (subject == null) {
                return findMailItems(olFolderName,
                    mail => mail.ReceivedTime.Date >= begin && mail.ReceivedTime.Date <= end);
            }
            return findMailItems(olFolderName, mail => mail.ReceivedTime.Date >= begin
                                                       && mail.ReceivedTime.Date <= end
                                                       && mail.Subject.Contains(subject));
        }

        public static IEnumerable<MailItem> findMailItemOn(string olFolderName, DateTime date, string subject = null) {
            if (subject == null) 
                return findMailItems(olFolderName, mail => mail.ReceivedTime.Date == date);
            
            return findMailItems(olFolderName, mail => mail.ReceivedTime.Date == date
                                                       && mail.Subject.Contains(subject));
        }

        private static IEnumerable<MailItem> findMailItems(string olFolderName, Predicate<MailItem> predicate) {
            var ns = application.GetNamespace("MAPI");
            var folder = ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Folders[olFolderName];
            return folder.Items.OfType<MailItem>().Where(mail => predicate(mail)).Select(mail => mail);
        }
    }
}