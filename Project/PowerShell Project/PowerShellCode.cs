using System;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace PowerShell_Project
{
    public class PowerShellCode
    {
        PowerShell ps = PowerShell.Create();
        public PowerShellCode()
        {

        }

        public void EOConnect()
        {

        }
        public Collection<PSObject> GetMailbox(Runspace runspace)
        {
            using (PowerShell shell = PowerShell.Create())
            {
                Console.WriteLine("Running : Get-Mailbox -ResultSize Unlimited");

                shell.Runspace = runspace;
                shell.AddScript("Get-Mailbox -ResultSize Unlimited");

                return shell.Invoke();
            }
        }
    }
}
