using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace AudioSwitcherProgram
{
    /// <summary>
    /// Wrapper for using PowerShell
    /// </summary>
    public static class PowerShellController
    {
        private static PowerShell powerShell;

        private const string _addOutputToStreamCommand = "Out-String";
        private const string _unknownErrorMessage = "Unknown Error";

        static PowerShellController()
        {
            powerShell = PowerShell.Create();
        }

        //Note: Not async as it could cause errors (multiple commands at the same time to the same PS terminal is not a good idea)
        /// <summary>
        /// Executes the provided script/command in PowerShell.
        /// </summary>
        /// <param name="script">The script/command to execute</param>
        /// <returns>The output of the command ran, or the error messages produced</returns>
        public static string SendCommand(string script)
        {
            string errorMsg = string.Empty;
            string output = string.Empty;

            powerShell.AddScript(script);

            //Adds the produced output to the stream so its readable by C#
            powerShell.AddCommand(_addOutputToStreamCommand);

            //Error handling
            powerShell.Streams.Error.DataAdded += (sender, eventArgs) =>
            {
                PSDataCollection<ErrorRecord>? errorCollection = sender as PSDataCollection<ErrorRecord>;

                if (errorCollection is null)
                {
                    //Unexpected sender of the event
                    //TODO: Add a clearer error message
                    errorMsg = _unknownErrorMessage;
                    return;
                }

                errorMsg = errorCollection[eventArgs.Index].ToString();
            };

            //Run the script
            PSDataCollection<PSObject> outputCollection = new PSDataCollection<PSObject>();
            
            //Note: Avoids making the method async and forces it to execute synchronously/in a blocking manner
            IAsyncResult asyncResult = powerShell.BeginInvoke<PSObject, PSObject>(null, outputCollection);
            powerShell.EndInvoke(asyncResult);

            //Check for errors
            if (!string.IsNullOrEmpty(errorMsg))
            {
                return errorMsg;
            }

            //Get output
            StringBuilder sb = new StringBuilder();

            foreach (PSObject outputItem in outputCollection)
            {
                sb.AppendLine(outputItem.BaseObject.ToString());
            }

            return sb.ToString().Trim();
        }

    }
}
