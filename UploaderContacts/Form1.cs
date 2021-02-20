using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Microsoft.Office.Interop.Outlook;

namespace UploaderContacts
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            var context = new ContactContext();
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
                        //sb.AppendLine($"{user.Name}\t {user.CompanyName} \t {user.Department} \t {user.BusinessTelephoneNumber} \t {user.JobTitle}");

                    }
                }
            }
            context.Save();
            textBox1.Text = sb.ToString();
        }
    }
}
