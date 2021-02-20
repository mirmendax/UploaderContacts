using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Microsoft.Office.Interop.Outlook;

namespace UploaderConsole
{
    class Program
    {
        public static void UploadAllContacts()
        {
            var context = new ContactContext("contacts.json");
            Application App = null;
            App = new Application();
            var gal = App.Session.GetGlobalAddressList();
            var sb = new StringBuilder();
            if (gal != null)
            {
                for (int i = 1; i < gal.AddressEntries.Count; i++)
                {
                    var item = gal.AddressEntries[i];
                    if (item.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry ||
                        item.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                    {
                        var user = item.GetExchangeUser();
                        context.Contacts.Add(new Contact
                        {
                            Id = Guid.NewGuid().ToString(),
                            Name = user.Name,
                            CompanyName = user.CompanyName,
                            Departament = user.Department,
                            JobTitle = user.JobTitle,
                            Phone = user.BusinessTelephoneNumber + ", " + user.MobileTelephoneNumber
                        });
                        Console.WriteLine($"{user.Name}\t {user.CompanyName} \t {user.Department} \t {user.JobTitle}");

                    }
                }
            }
            context.Save();
            Console.WriteLine();
            Console.WriteLine("===================END=======================");
            Console.ReadLine();
        }

        public static void SelectionContactsVolGes()
        {
            var newcontact = new ContactContext("volGES.json");
            var allContacts = new ContactContext("contacts.json");
            allContacts.Load();
            foreach (var contact in allContacts.Contacts)
            {
                if (!string.IsNullOrWhiteSpace(contact.CompanyName))
                {
                    if (contact.CompanyName.Contains("Волжская ГЭС"))
                    {
                        newcontact.Contacts.Add(contact);
                        Console.WriteLine($"Add {contact.Name}");
                    }
                }
            }
            newcontact.Save();
        }
        public static void SelectionContactsVKK()
        {
            var newcontact = new ContactContext("gidroremontVKK.json");
            var allContacts = new ContactContext("contacts.json");
            allContacts.Load();
            foreach (var contact in allContacts.Contacts)
            {
                if (!string.IsNullOrWhiteSpace(contact.Departament))
                {
                    if (contact.Departament.Contains("Волжский"))
                    {
                        newcontact.Contacts.Add(contact);
                        Console.WriteLine($"Add {contact.Name}");
                    }
                }
            }
            newcontact.Save();
        }

        static void Main(string[] args)
        {
            SelectionContactsVKK();
            Console.ReadLine();
        }
    }
}
