//
// vCard	 https://tools.ietf.org/html/rfc6350
// vCalendar https://datatracker.ietf.org/doc/html/rfc5545
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
		static bool forwardOnlyCalendars = true;

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
			Outlook.MAPIFolder calendar = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

			// args
			var tmp = args.Arg("ignore");

			if(tmp != null)
				ignoreCat = tmp.Split(',').Select(x => x.Trim());

			devnull = args.ArgBool("devnull");
			phoneOnly = args.ArgBool("phoneonly");
			forwardOnlyCalendars = args.ArgBool("forward");

			PrintContacts(contacts);
			PrintCalendar(calendar);
		}

		static void PrintContacts(Outlook.MAPIFolder contacts)
		{
			const string dirName = "Contacts";

			if (!Directory.Exists(dirName))
				Directory.CreateDirectory(dirName);

			foreach (Outlook.MAPIFolder folder in contacts.Folders)
			{
				if (folder.Name.Equals("archive", StringComparison.CurrentCultureIgnoreCase))
					continue;

				PrintContacts(folder);
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
				  : new StreamWriter(Path.Combine(dirName, contact.FullName.StripIllegalChars() + ".vcf"));

				file.WriteLine(
$@"BEGIN:VCARD
VERSION:3.0
N:{contact.LastName};{contact.FirstName}
FN:{contact.FullName}
TEL;work;voice:{contact.BusinessTelephoneNumber}
TEL;home;voice:{contact.HomeTelephoneNumber}
TEL;cell;voice:{contact.MobileTelephoneNumber}
ADR;work:;;{contact.BusinessAddress?.Replace(Environment.NewLine, ";")}
ADR;home:;;{contact.HomeAddress?.Replace(Environment.NewLine, ";")}
EMAIL:{contact.Email1Address}
EMAIL2:{contact.Email2Address}
EMAIL3:{contact.Email3Address}
NOTE:{contact.Body?.Replace(Environment.NewLine, ";")}
CATEGORIES:{contact.Categories}
URL:{contact.WebPage}
REV:{contact.LastModificationTime:yyyyMMddThhmmssZ}
END:VCARD");

				file.Close();
			}
		}

		static void PrintCalendar(Outlook.MAPIFolder calanders)
		{
			const string dirName = "Calendar";

			if (!Directory.Exists(dirName))
				Directory.CreateDirectory(dirName);

			foreach (Outlook.MAPIFolder folder in calanders.Folders)
			{
				if (folder.Name.Equals("archive", StringComparison.CurrentCultureIgnoreCase))
					continue;

				PrintCalendar(folder);
			}

			foreach (var obj in calanders.Items)
			{
				Outlook.AppointmentItem calendar = obj as Outlook.AppointmentItem;

				if (calendar == null)
					continue;

				if (calendar.Start < DateTime.Now && !calendar.IsRecurring)
					continue;

				Console.WriteLine(calendar.Subject);

				var file = devnull ?
				new StreamWriter(Stream.Null)
			  : new StreamWriter(Path.Combine(dirName, calendar.Subject.StripIllegalChars() + ".vcs"));

				file.WriteLine(
$@"BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//OutlookExport v1.0//EN
BEGIN:VEVENT
DTSTAMP:{calendar.CreationTime:yyyyMMddThhmmssZ}
DTSTART:{calendar.Start:yyyyMMddThhmmssZ}
DTEND:{calendar.End:yyyyMMddThhmmssZ}
CLASS:{calendar.Sensitivity}
CATEGORIES:{calendar.Categories}
SUMMARY:{calendar.Subject}
END:VEVENT
END:VCALENDAR");
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

	static class StringExtensions
	{
		internal static string StripIllegalChars(this string name)
		{
			char[] illegal = new char[] { '\\','/',':','*', '?', '"','<','>','|' };

			string retValue = name;

			foreach(var c in illegal)
				retValue = retValue.Replace(c, '_');

			return retValue;
		}
	}
}
