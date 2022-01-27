using System;
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
            PSCommand sc;
            Collection<PSObject> results;

            using (Runspace runspace = RunspaceFactory.CreateRunspace(RunspaceConfiguration.Create()))
            {
                // Open local runspace
                runspace.Open();

                // Load Exchange Online PowerShell V2 module
                using (PowerShell shell = PowerShell.Create())
                {
                    sc = shell.Commands;
                    Console.WriteLine("Setting: Execution Policy");
                    shell.Runspace = runspace;
                    shell.Commands.AddCommand("Set-ExecutionPolicy");
                    shell.Commands.AddParameter("ExecutionPolicy", "RemoteSigned");
                    shell.Commands.AddParameter("Scope", "CurrentUser");
                    shell.Invoke();
                    Console.WriteLine("Running : Import-Module ExchangeOnlineManagement");
                    shell.Runspace = runspace;
                    sc.AddCommand("Import-Module");
                    sc.AddParameter("Name", "ExchangeOnlineManagement");
                    shell.Invoke();
                    
                    WriteStreams(shell.Streams);
                }

                // Connect to Exchange Online
                ConnectExchangeOnlineUsingUserPrinsipalNameAndPassword(runspace);
                // Run Get-Mailbox
                results = psc.GetMailbox(runspace);
                fillMailboxes(results);
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

        private void fillMailboxes(Collection<PSObject> results)
        {
            foreach (var item in results)
            {
                listBox1.Items.Add(item.Properties["DisplayName"].Value.ToString());
            }
            
        }

        private void ConnectExchangeOnlineUsingUserPrinsipalNameAndPassword(Runspace runspace)
        {
            PSCommand sc;
            using (PowerShell shell = PowerShell.Create())
            {
                sc = shell.Commands;
                Console.WriteLine("Running : Connect-ExchangeOnline with AppID");

                shell.Runspace = runspace;
                sc.AddCommand("Connect-ExchangeOnline");
                sc.AddParameter("AppId", "8c4e320e-b0a2-4958-9d5d-17275b3be6ba");
                sc.AddParameter("CertificateThumbprint", "64BC31A8FFFC7C86146448171784E6C0BE3F7CE4");
                sc.AddParameter("Organization", "nrz5.onmicrosoft.com");
                sc.AddParameter("ShowProgress", false);
                sc.AddParameter("ShowBanner", false);
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

        private void button2_Click(object sender, EventArgs e)
        {
            installModule();
        }

        private void installModule()
        {
            // Load Exchange Online PowerShell V2 module
            using (PowerShell shell = PowerShell.Create())
            {
                Console.WriteLine("Installing NuGet");
                shell.Commands.AddCommand("Install-PackageProvider");
                shell.Commands.AddParameter("Name", "NuGet");
                shell.Commands.AddParameter("Confirm");
                shell.Commands.AddParameter("Scope", "CurrentUser");
                shell.Invoke();

                Console.WriteLine("Installing Module ExchangeOnlineManagement");
                shell.Commands.AddCommand("Install-Module");
                shell.Commands.AddParameter("Name", "ExchangeOnlineMangement");
                shell.Commands.AddParameter("Confirm");
                shell.Commands.AddParameter("Scope", "CurrentUser");
                shell.Invoke();
                WriteStreams(shell.Streams);
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Console.WriteLine(listBox1.SelectedItem.ToString());
        }
    }
}
