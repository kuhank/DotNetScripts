#region Help:  Introduction to the script task
/* The Script Task allows you to perform virtually any operation that can be accomplished in
 * a .Net application within the context of an Integration Services control flow. 
 * 
 * Expand the other regions which have "Help" prefixes for examples of specific ways to use
 * Integration Services features within this script task. */
#endregion


#region Namespaces
using System;
using System.Data;
using Microsoft.SqlServer.Dts.Runtime;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Net.Mail;
#endregion

namespace ST_aeee1c9a543e4821b11bd0d20a3dda2c
{
    /// <summary>
    /// ScriptMain is the entry point class of the script.  Do not change the name, attributes,
    /// or parent of this class.
    /// </summary>
	[Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute]
	public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
	{
        #region Help:  Using Integration Services variables and parameters in a script
        /* To use a variable in this script, first ensure that the variable has been added to 
         * either the list contained in the ReadOnlyVariables property or the list contained in 
         * the ReadWriteVariables property of this script task, according to whether or not your
         * code needs to write to the variable.  To add the variable, save this script, close this instance of
         * Visual Studio, and update the ReadOnlyVariables and 
         * ReadWriteVariables properties in the Script Transformation Editor window.
         * To use a parameter in this script, follow the same steps. Parameters are always read-only.
         * 
         * Example of reading from a variable:
         *  DateTime startTime = (DateTime) Dts.Variables["System::StartTime"].Value;
         * 
         * Example of writing to a variable:
         *  Dts.Variables["User::myStringVariable"].Value = "new value";
         * 
         * Example of reading from a package parameter:
         *  int batchId = (int) Dts.Variables["$Package::batchId"].Value;
         *  
         * Example of reading from a project parameter:
         *  int batchId = (int) Dts.Variables["$Project::batchId"].Value;
         * 
         * Example of reading from a sensitive project parameter:
         *  int batchId = (int) Dts.Variables["$Project::batchId"].GetSensitiveValue();
         * */

        #endregion

        #region Help:  Firing Integration Services events from a script
        /* This script task can fire events for logging purposes.
         * 
         * Example of firing an error event:
         *  Dts.Events.FireError(18, "Process Values", "Bad value", "", 0);
         * 
         * Example of firing an information event:
         *  Dts.Events.FireInformation(3, "Process Values", "Processing has started", "", 0, ref fireAgain)
         * 
         * Example of firing a warning event:
         *  Dts.Events.FireWarning(14, "Process Values", "No values received for input", "", 0);
         * */
        #endregion

        #region Help:  Using Integration Services connection managers in a script
        /* Some types of connection managers can be used in this script task.  See the topic 
         * "Working with Connection Managers Programatically" for details.
         * 
         * Example of using an ADO.Net connection manager:
         *  object rawConnection = Dts.Connections["Sales DB"].AcquireConnection(Dts.Transaction);
         *  SqlConnection myADONETConnection = (SqlConnection)rawConnection;
         *  //Use the connection in some code here, then release the connection
         *  Dts.Connections["Sales DB"].ReleaseConnection(rawConnection);
         *
         * Example of using a File connection manager
         *  object rawConnection = Dts.Connections["Prices.zip"].AcquireConnection(Dts.Transaction);
         *  string filePath = (string)rawConnection;
         *  //Use the connection in some code here, then release the connection
         *  Dts.Connections["Prices.zip"].ReleaseConnection(rawConnection);
         * */
        #endregion


		/// <summary>
        /// This method is called when this script task executes in the control flow.
        /// Before returning from this method, set the value of Dts.TaskResult to indicate success or failure.
        /// To open Help, press F1.
        /// </summary>
		public void Main()
		{
            try
            {
                const string pass = "Password";
				
                var toadd1 = Dts.Variables["toadd1"].Value.ToString();
                var toaddname1 = Dts.Variables["toaddname1"].Value.ToString();
                var toadd2 = Dts.Variables["toadd2"].Value.ToString();
                var toaddname2 = Dts.Variables["toaddname2"].Value.ToString();
                var ccadd1 = Dts.Variables["ccadd1"].Value.ToString();
                var ccadd2 = Dts.Variables["ccadd2"].Value.ToString();
                var ccaddname1= Dts.Variables["ccaddname1"].Value.ToString();
                var ccaddname2= Dts.Variables["ccaddname2"].Value.ToString();
				var portno= Dts.Variables["portnumber"].Value.ToString();
				
				
                var fromaddress = new MailAddress("noreply@domain.com", "NoReply");
				
                var toaddress1 = new MailAddress(toadd1, toaddname1);
                var toaddress2 = new MailAddress(toadd2, toaddname2);
                var ccaddress1 = new MailAddress(ccadd1, ccaddname1);
                var ccaddress2 = new MailAddress(ccadd2, ccaddname2);

                MailMessage emailMessage = new MailMessage();
                emailMessage.From = new MailAddress(fromaddress.ToString());
                emailMessage.To.Add(toaddress1);
                emailMessage.CC.Add(ccaddress1);
                emailMessage.To.Add(toaddress2);
                emailMessage.CC.Add(ccaddress2);
                emailMessage.Subject = Dts.Variables["Subject"].Value.ToString();
                emailMessage.Body = Dts.Variables["Body"].Value.ToString();
                emailMessage.Priority = MailPriority.Normal;
                SmtpClient MailClient = new SmtpClient("smtp.office365.com", Portno)
                {
                    Host = "smtp.office365.com",
                    Port = portno,
                    UseDefaultCredentials = false,
                    EnableSsl = true
                };
                MailClient.UseDefaultCredentials = false;
                MailClient.EnableSsl = true;
                MailClient.Port = portno;
                MailClient.Host = "smtp.office365.com";
                MailClient.Credentials = new System.Net.NetworkCredential(fromaddress.Address, pass);

                MailClient.TargetName = "STARTTLS/smtp.office365.com";
                MailClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                MailClient.Send(emailMessage);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            Dts.TaskResult = (int)ScriptResults.Success;
		}

        #region ScriptResults declaration
        /// <summary>
        /// This enum provides a convenient shorthand within the scope of this class for setting the
        /// result of the script.
        /// 
        /// This code was generated automatically.
        /// </summary>
        enum ScriptResults
        {
            Success = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success,
            Failure = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure
        };
        #endregion

	}
}