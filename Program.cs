//
// https://tools.ietf.org/html/rfc6350
//
//

using System;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookExport
{
	class Program
	{
		static void Main(string[] args)
		{
			Outlook.Application outlook = null;

			try
			{
				outlook = (Outlook.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application");
			}
			catch(Exception)
			{
				outlook = new Outlook.Application();
			}

			var ns = outlook.GetNamespace("MAPI");

			Outlook.MAPIFolder contacts = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);

			var path = args.Length == 1 ? args[0] : "";

			if (!Directory.Exists(path))
				Directory.CreateDirectory(path);

			PrintContacts(contacts, path);
		}

		static void PrintContacts(Outlook.MAPIFolder contacts, string path)
		{
			foreach (Outlook.MAPIFolder folder in contacts.Folders)
			{
				if (folder.Name.Equals("archive", StringComparison.CurrentCultureIgnoreCase))
					continue;

				PrintContacts(folder, path);
			}

			foreach (var obj in contacts.Items)
			{
				Outlook.ContactItem contact = obj as Outlook.ContactItem;

				if (contact == null)
					continue;

				if (string.IsNullOrEmpty(contact.BusinessTelephoneNumber)
				&& string.IsNullOrEmpty(contact.MobileTelephoneNumber)
				&& string.IsNullOrEmpty(contact.MobileTelephoneNumber))
					continue;

				Console.WriteLine(contact.FullName);

				var file = new StreamWriter(Path.Combine(path, contact.FullName.Replace('/',',') + ".vcf"));

				file.WriteLine(
$@"BEGIN:VCARD
VERSION:3.0
N: {contact.LastName};{contact.FirstName}
FN: {contact.FullName}
TEL;work;voice:{contact.BusinessTelephoneNumber}
TEL;home;voice:{contact.HomeTelephoneNumber}
TEL;cell;voice:{contact.MobileTelephoneNumber}
ADR;work:;;{contact.BusinessAddress?.Replace(Environment.NewLine, ";")}
ADR;home:;;{contact.HomeAddress?.Replace(Environment.NewLine, ";")}
EMAIL:{contact.Email1Address}
EMAIL2:{contact.Email2Address}
EMAIL3:{contact.Email3Address}
NOTE:{contact.Body?.Replace(Environment.NewLine, ";")}
URL:{contact.WebPage}
REV:{contact.LastModificationTime.ToString("yyyyMMddThhmmssZ")}
END:VCARD");

				file.Close();
			}
		}
	}
}
