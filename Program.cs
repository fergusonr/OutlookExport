//
// https://tools.ietf.org/html/rfc6350
//
//

using System;
using System.Linq;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;

namespace OutlookExport
{
	static class Program
	{
		static IEnumerable<string> ignoreCat;
		static bool devnull;
		static bool phoneOnly;

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

			// args
			var path = args.Arg("path");
			var tmp = args.Arg("ignore");

			if(tmp != null)
				ignoreCat = tmp.Split(',').Select(x => x.Trim());

			devnull = args.ArgBool("devnull");
			phoneOnly = args.ArgBool("phoneonly");

			if (path != null && !Directory.Exists(path))
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

				if (phoneOnly
				&& string.IsNullOrEmpty(contact.BusinessTelephoneNumber)
				&& string.IsNullOrEmpty(contact.HomeTelephoneNumber)
				&& string.IsNullOrEmpty(contact.MobileTelephoneNumber))
					continue;

				if (ignoreCat != null && contact.Categories != null)
				{
					var cats = contact.Categories.Split(',').Select(x => x.Trim());

					if (cats.Intersect(ignoreCat, EqualityComparer<string>.Default).Any())
						continue;
				}

				Console.WriteLine(contact.FullName);

				var file = devnull ? 
					new StreamWriter(Stream.Null) 
				  : new StreamWriter(Path.Combine(path, contact.FullName.Replace('/',',') + ".vcf"));

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

	static class ArgsParse
	{
		internal static string Arg(this string[] args, string name)
		{
			var index = Array.IndexOf(args, $"-{name}");
			return index != -1 && index + 1 < args.Length ? args[index + 1] : null;
		}

		internal static bool ArgBool(this string[] args, string name)
		{
			var index = Array.IndexOf(args, $"-{name}");
			return index != -1 ? true : false;
		}

	}
}
