using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Windows.Forms;

namespace PowerShell_Project
{
    public partial class Form1 : Form
    {
        PowerShellCode psc;

        public Form1()
        {
            InitializeComponent();
            psc = new PowerShellCode();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void ConnectEO()
        {
            PSCommand sc;
            Collection<PSObject> results;

            using (Runspace runspace = RunspaceFactory.CreateRunspace(RunspaceConfiguration.Create()))
            {
                // Run Get-Mailbox

                // Disconnect from Exchange Online
            }
        }

        private void fillList()
        {
            foreach (var item in psc.GetMailbox())
            {
                listBox1.Items.Add(item.Properties["DisplayName"].Value.ToString());
            }
            
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            psc.DisconnectFromExO();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            psc.EOConnect();
            fillList();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            psc.installModule();
        }



        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Console.WriteLine(listBox1.SelectedItem.ToString());
            proxyEmail();
        }
        private void proxyEmail()
        {
            listBox2.Items.Clear();
            PSObject mailAdresses = new PSObject();
            int i = 0;
            foreach (var item in psc.getProxyMail(listBox1.SelectedItem.ToString()))
            {
                mailAdresses = item.Properties["EmailAddresses"].Value.ToString();
            }
            string[] splitted = mailAdresses.Properties["EmailAddresses"].Value.ToString().Split(' ');
            foreach (var item in splitted)
            {
                listBox2.Items.Add(item);
            }

        }
    }
}
