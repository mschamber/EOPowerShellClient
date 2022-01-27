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
using System.Management.Automation.Runspaces;
using System.Management.Automation.Remoting;

namespace PowerShell_Project
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }
        public static void WriteStreams(PSDataStreams Streams)
        {
            // Write Error, Warning and Information stream

            PSDataCollection<ErrorRecord> errorStream = Streams.Error;
            foreach (ErrorRecord errorRecord in errorStream)
            {
                Console.WriteLine(errorRecord.ToString());
            }

            PSDataCollection<WarningRecord> warningRecords = Streams.Warning;
            foreach (WarningRecord warningRecord in warningRecords)
            {
                Console.WriteLine(warningRecord.ToString());
            }

            PSDataCollection<InformationRecord> informationRecords = Streams.Information;
            foreach (InformationRecord informationRecord in informationRecords)
            {
                Console.WriteLine(informationRecord.ToString());
            }

            Console.WriteLine("");
        }
        private void ConnectEO()
        {

            Collection<PSObject> results;

            using (Runspace runspace = RunspaceFactory.CreateRunspace(RunspaceConfiguration.Create()))
            {
                // Open local runspace
                runspace.Open();

                // Load Exchange Online PowerShell V2 module
                using (PowerShell shell = PowerShell.Create())
                {
                    Console.WriteLine("Running : Import-Module ExchangeOnlineManagement");

                    shell.Runspace = runspace;
                    shell.Commands.AddScript("Import-Module ExchangeOnlineManagement");
                    shell.Invoke();

                    WriteStreams(shell.Streams);
                }

                // Connect to Exchange Online
                ConnectExchangeOnlineUsingUserPrinsipalNameAndPassword(runspace);
                //ConnectExchangeOnlineUsingPfxFile(runspace);
                //ConnectExchangeOnlineUsingThumbprint(runspace);
                //ConnectExchangeOnlineUsingX509Certificate2(runspace);

                // Run Get-Mailbox
                using (PowerShell shell = PowerShell.Create())
                {
                    Console.WriteLine("Running : Get-Mailbox -ResultSize Unlimited");

                    shell.Runspace = runspace;
                    shell.AddScript("Get-Mailbox -ResultSize Unlimited");

                    results = shell.Invoke();

                    foreach (PSObject result in results)
                    {
                        Console.WriteLine(result.Properties["UserPrincipalName"].Value.ToString());
                    }

                    WriteStreams(shell.Streams);
                }
                // Disconnect from Exchange Online
                using (PowerShell shell = PowerShell.Create())
                {
                    Console.WriteLine("Running : Disconnect-ExchangeOnline -Confirm:$false");

                    shell.Runspace = runspace;
                    shell.AddScript("Disconnect-ExchangeOnline -Confirm:$false");

                    results = shell.Invoke();
                    foreach (PSObject result in results)
                    {
                        Console.WriteLine(result.Properties["UserPrincipalName"].Value.ToString());
                    }

                    WriteStreams(shell.Streams);
                }
            }
        }

        private void ConnectExchangeOnlineUsingUserPrinsipalNameAndPassword(Runspace runspace)
        {
            using (PowerShell shell = PowerShell.Create())
            {
                Console.WriteLine("Running : Connect-ExchangeOnline -Credential <Cred> -ShowProgress $false -ShowBanner $false");

                string userPrincipalName = "michael.schamber@nrz5.onmicrosoft.com";
                string password = "Dpskdsb5771!";

                SecureString secureString = new SecureString();

                foreach (var ch in password)
                {
                    secureString.AppendChar(ch);
                }

                PSCredential pSCredential = new PSCredential(userPrincipalName, secureString);

                shell.Runspace = runspace;
                shell.Commands.AddCommand("Connect-ExchangeOnline");
                shell.Commands.AddParameter("AppId", "8c4e320e-b0a2-4958-9d5d-17275b3be6ba");
                shell.Commands.AddParameter("CertificateThumbprint", "64BC31A8FFFC7C86146448171784E6C0BE3F7CE4");
                shell.Commands.AddParameter("Organization", "nrz5.onmicrosoft.com");
                shell.Commands.AddParameter("ShowProgress", false);
                shell.Commands.AddParameter("ShowBanner", false);
                // shell.Commands.AddParameter("PSSessionOption", new PSSessionOption() { ProxyAccessType = ProxyAccessType.IEConfig });
                //Connect-ExchangeOnline -AppId 8c4e320e-b0a2-4958-9d5d-17275b3be6ba -CertificateThumbprint "64BC31A8FFFC7C86146448171784E6C0BE3F7CE4" -Organization nrz5.onmicrosoft.com
                shell.Invoke();

                WriteStreams(shell.Streams);
            }
        }

        private void DisconnectEO()
        {
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
