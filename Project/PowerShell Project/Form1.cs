using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management.Automation;
using System.Collections.ObjectModel;
using System.Security;

namespace PowerShell_Project
{
    public partial class Form1 : Form
    {
        PowerShell ps;
        public Form1()
        {
            InitializeComponent();
            ps = PowerShell.Create();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void ConnectEO()
        {
            string userName = "michael.schamber@nrz5.onmicrosoft.com";
            string password = "Dpskdsb5771!";
            SecureString secureString = new SecureString();
            foreach (var item in password)
            {
                secureString.AppendChar(item);
            }
            secureString.MakeReadOnly();

            PSCredential credential = new PSCredential(userName, secureString);

            ps.AddCommand("Set-ExecutionPolicy")
                .AddParameter("ExecutionPolicy","RemoteSigned")
                .AddParameter("Scope","LocalMachine")
                .Invoke();
            ps.AddCommand("Import-Module").AddParameter("Name","ExchangeOnlineManagement").Invoke();
            var result = ps.AddCommand("Connect-ExchangeOnline").Invoke();
            var mailboxes = ps.AddCommand("Get-Mailbox").Invoke();

        }
        private void DisconnectEO()
        {
            ps.AddCommand("Disconnect-ExchangeOnline").Invoke();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DisconnectEO();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ConnectEO();
        }
    }
}
