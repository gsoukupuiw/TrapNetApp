using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TrapNetPluginTest2
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var inspector = Globals.ThisAddIn.Application.ActiveInspector();
            if (inspector != null && inspector.CurrentItem is Microsoft.Office.Interop.Outlook.MailItem mailItem)
            {
                string emailBody = mailItem.Body;

                // Truncate the email body to fit model requirements
                if (emailBody.Length > 512)
                {
                    emailBody = emailBody.Substring(0, 512);
                }

                // Set up the process start info using the specific path to your Python script
                var psi = new ProcessStartInfo
                {
                    FileName = @"C:\Users\Admin\source\repos\TrapNetPluginTest1\TrapNetPluginTest1", // Assuming Python is in your system's PATH
                    Arguments = @"C:\Users\Admin\source\repos\TrapNetPluginTest1\TrapNetPluginTest1\TrapNetPluginTest1.py",
                    UseShellExecute = false,
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true, // Capture standard error
                    CreateNoWindow = true
                };

                // Execute the Python script and pass the email body via standard input
                using (var process = Process.Start(psi))
                {
                    process.StandardInput.WriteLine(emailBody);
                    process.StandardInput.Close(); // Ensure to close the input stream
                    string output = process.StandardOutput.ReadToEnd(); // Capture the script output
                    string errors = process.StandardError.ReadToEnd(); // Capture any errors

                    // Display script output or errors
                    if (!string.IsNullOrEmpty(errors))
                    {
                        MessageBox.Show(errors, "Script Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        MessageBox.Show(output, "Script Output", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else
            {
                MessageBox.Show("Please select an email.", "No Email Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

    }
}
