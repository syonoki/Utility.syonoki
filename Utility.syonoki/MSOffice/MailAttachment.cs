using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Outlook;

namespace Utility.syonoki.MSOffice {
    public static class MailAttachment {
        #region from mail item

        public static IEnumerable<string> attachmentList(string olFolderName, string containedSubject) {
            MailItem findedItem;
            bool hasMailFounded = OutlookSimpleApi.hasMailReceived(olFolderName, containedSubject, out findedItem);

            if (hasMailFounded)
                for (int i = 0; i < findedItem.Attachments.Count; i++)
                    yield return findedItem.Attachments[i].FileName;
        }

        public static bool findAttachedFile(MailItem mail, string partOfAttachmentName) {
            return mail.Attachments.OfType<Attachment>()
                .Any(x => x.FileName.Contains(partOfAttachmentName));
        }

        public static string saveAttachedFile(MailItem mail, string partOfAttachmentName, string saveDirectory) {
            var attachment = mail.Attachments.OfType<Attachment>()
                .SingleOrDefault(x => x.FileName.Contains(partOfAttachmentName)
                                      && Path.GetExtension(x.FileName).Contains("xl"));

            if (attachment == null)
                return null;

            return saveAttachedFiles(saveDirectory, attachment);
        }

        #endregion

        #region from specific Date

        public static void saveAttachedFilesAtSpecificDate(DateTime date, string olFolderName, string saveDirectory) {
            if (olFolderName == null) throw new ArgumentNullException(paramName: nameof(olFolderName));

            MAPIFolder olFolder = OutlookSimpleApi.getMailFolder(olFolderName);

            List<MailItem> mailsAtDate = (from m in olFolder.Items.OfType<MailItem>()
                where m.ReceivedTime.Date == date.Date
                select m).ToList();

            mailsAtDate.ForEach(m => {
                var attachment = m.Attachments.OfType<Attachment>();
                attachment.ToList().ForEach(a => { saveAttachedFiles(saveDirectory, a); });
            });
        }

        private static string saveAttachedFiles(string saveDirectory, Attachment attachment, bool overwrite = true) {
            string path = $@"{saveDirectory}\{attachment.FileName}";
            if (File.Exists(path)) {
                if (overwrite) File.Delete(path);
                else return null;
            }

            attachment.SaveAsFile(path);
            return path;
        }

        #endregion
    }

    public static class MailAttachmentExtension {
        public static bool findAttachedFile(this MailItem mail, string partOfAttachmentName)
            => MailAttachment.findAttachedFile(mail, partOfAttachmentName);

        public static string saveAttachedFile(this MailItem mail, string partOfAttachmentName, string saveDirectory)
            => MailAttachment.saveAttachedFile(mail, partOfAttachmentName, saveDirectory);
    }
}