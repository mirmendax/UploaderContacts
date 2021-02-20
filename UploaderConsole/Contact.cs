using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;

namespace UploaderConsole
{
    public class Contact
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string CompanyName { get; set; }
        public string Departament { get; set; }
        public string JobTitle { get; set; }
        public string Phone { get; set; }
    }

    public class ContactContext
    {
        public List<Contact> Contacts = new List<Contact>();
        private readonly string FILE = "contacts.json";

        public ContactContext(string file)
        {
            FILE = file;
        }

        public void Save()
        {
            var json = JsonConvert.SerializeObject(Contacts, Formatting.Indented);
            File.WriteAllText(FILE, json);
        }

        public void Load()
        {
            if (File.Exists(FILE)) 
            {
                var json = File.ReadAllText(FILE);
                Contacts = JsonConvert.DeserializeObject<List<Contact>>(json);
            }
        }
    }
}
