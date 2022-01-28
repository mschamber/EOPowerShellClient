using System;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace PowerShell_Project
{
    public class PowerShellCode
    {
        Runspace runspace;
        PowerShell ps;
        PSCommand sc;
        public PowerShellCode()
        {
            runspace = RunspaceFactory.CreateRunspace(RunspaceConfiguration.Create());
            ps = PowerShell.Create();
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
        public void installModule()
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
        public void EOConnect()
        {
            runspace.Open();

            // Load Exchange Online PowerShell V2 module
            using (PowerShell shell = PowerShell.Create())
            {
                sc = shell.Commands;
                Console.WriteLine("Setting: Execution Policy");
                shell.Runspace = runspace;
                sc.AddCommand("Set-ExecutionPolicy");
                sc.AddParameter("ExecutionPolicy", "RemoteSigned");
                sc.AddParameter("Scope", "CurrentUser");
                shell.Invoke();
                Console.WriteLine("Running : Import-Module ExchangeOnlineManagement");
                shell.Runspace = runspace;
                sc.AddCommand("Import-Module");
                sc.AddParameter("Name", "ExchangeOnlineManagement");
                shell.Invoke();

                WriteStreams(shell.Streams);

                ConnectExchangeOnlineUsingUserPrinsipalNameAndPassword(runspace);
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
        public void DisconnectFromExO()
        {
            using (PowerShell shell = PowerShell.Create())
            {
                Console.WriteLine("Running : Disconnect-ExchangeOnline -Confirm:$false");

                shell.Runspace = runspace;
                shell.AddScript("Disconnect-ExchangeOnline -Confirm:$false");

                shell.Invoke();
                WriteStreams(shell.Streams);
            }
        }
        public Collection<PSObject> GetMailbox()
        {
            using (PowerShell shell = PowerShell.Create())
            {
                Console.WriteLine("Running : Get-Mailbox -ResultSize Unlimited");

                shell.Runspace = runspace;
                shell.AddScript("Get-Mailbox -ResultSize Unlimited");

                return shell.Invoke();
            }
        }
        public Collection<PSObject> getProxyMail(string DisplayName)
        {
            using (PowerShell shell = PowerShell.Create())
            {
                Console.WriteLine("Running : Get-Mailbox -ResultSize Unlimited");

                shell.Runspace = runspace;
                shell.AddScript("Get-Mailbox '"+DisplayName+"' -ResultSize Unlimited");

                return shell.Invoke();
            }
        }

    }
}
