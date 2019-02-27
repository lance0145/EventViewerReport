using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net;
using System.IO;

namespace EventViewerReport
{
    class Program
    {
       
        static void Main(string[] args)
        {
            RecordEvents();
        }

        static void RecordEvents()
        {
            EventLogHelper evnt = new EventLogHelper();
            Utility util = new Utility();
            string queryResult = "";           
            List<string> ErrorsEvents = new List<string>();
            List<string> SuccessEvents = new List<string>();
            List<string> ErrorsEventsConsole = new List<string>();
            List<string> SuccessEventsConsole = new List<string>();
            List<string> EmailHtmlEvents = new List<string>();
            List<string> IpAddressList = new List<string>();
            var currentDirectory = Directory.GetCurrentDirectory();
            File.WriteAllText(currentDirectory + "\\ErrorEvents.txt", String.Empty);
            string[] IpAddressArray = File.ReadAllLines(currentDirectory + "\\IPSToCheck.txt");
            DateTime StartTime = DateTime.Now;
            TimeZone LocalZone = TimeZone.CurrentTimeZone;
            Console.WriteLine("Server error events as of {0} at {1} ({2})", StartTime.ToString("yyyy/MM/dd"), StartTime.ToString("hh:mm:ss tt"), LocalZone.StandardName);
            EmailHtmlEvents.Add("<p>Server error events as of " + StartTime.ToString("yyyy/MM/dd")+ " at " + StartTime.ToString("hh:mm:ss tt") + " (" + LocalZone.StandardName + ")" + "</p>");

            try
            {
                
                if (IpAddressArray != null && IpAddressArray.Length > 0)
                {

                    for (int i = 0; i < IpAddressArray.Length; i++)
                    {
                                           
                        try
                        {
                            if (!string.IsNullOrEmpty(IpAddressArray[i]))
                            {
                                if (ValidateIPAddress(IpAddressArray[i]))
                                {
                                    Console.WriteLine("Checking.... {0}", IpAddressArray[i]);
                                    IpAddressList.Add(IpAddressArray[i]);
                                    try
                                    {
                                        queryResult = evnt.GetRecentEvents(IpAddressArray[i]);
                                        if (queryResult == "        No Events Error Found in this Machine.")
                                        {
                                            SuccessEvents.Add("<li>" + IpAddressArray[i] + "<ul>" + queryResult + "</ul></li>");
                                            SuccessEventsConsole.Add(IpAddressArray[i] + "\r\n" + queryResult.Replace("<li>", "").Replace("</li>", ""));
                                        }
                                        else
                                        { 
                                            ErrorsEvents.Add("<li>" + IpAddressArray[i] + "<ul>" + queryResult + "</ul></li>");
                                            ErrorsEventsConsole.Add(IpAddressArray[i] + "\r\n" + queryResult.Replace("<li>", "").Replace("</li>", ""));
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        //  requires elevation of privilege 
                                        ErrorsEvents.Add("<li>" + IpAddressArray[i] + "<ul><li>" + "ERROR: Could not read Event Log, Check Firewall and Admin Account" + "</li></ul></li>");
                                        ErrorsEventsConsole.Add(IpAddressArray[i] + "\r\n" + "        ERROR: Could not read Event Log, Check Firewall and Admin Account");
                                    }                            
                                 
                                }
                            }
                        }
                        catch { }
                    }
                    if (ErrorsEvents.Count != 0)
                    {
                        Console.WriteLine("\r\nERRORS\r\n-------------------------------------------------------------------------");
                        EmailHtmlEvents.Add("<br><h2><font style=\"color: red\">ERRORS<hr></font></h2>");
                        Console.WriteLine(string.Join("\r\n", ErrorsEventsConsole));
                        EmailHtmlEvents.Add("<ul>" + string.Join("\r\n", ErrorsEvents) + "</ul>");
                    }
                    else
                    {
                        Console.WriteLine("No ERRORS Found on Servers.");
                        EmailHtmlEvents.Add("<p>No ERRORS Found on Servers.</p>");

                    }
                    if (SuccessEvents.Count != 0)
                    {
                        Console.WriteLine("\r\nSUCCESS\r\n-------------------------------------------------------------------------");
                        EmailHtmlEvents.Add("<br><h2><font style=\"color: green\">SUCCESS<hr></font></h2>");
                        Console.WriteLine(string.Join("\r\n", SuccessEventsConsole));
                        EmailHtmlEvents.Add("<ul>" + string.Join("\r\n", SuccessEvents) + "</ul>");
                    }
                    DateTime EndTime = DateTime.Now;
                    Console.WriteLine("\r\nFinish at {0}", EndTime.ToString("hh:mm:ss tt"));                    
                    EmailHtmlEvents.Add("<br><p>Finish at " + EndTime.ToString("hh:mm:ss tt") + "</p>");
                    Console.WriteLine("Servers checked include: " + string.Join(" | ", IpAddressList));
                    EmailHtmlEvents.Add("<p>Servers checked include: "+ string.Join(" | ", IpAddressList) + "</p>");
                    util.WriteToFile("ErrorEvents.txt", string.Join("\r\n" , EmailHtmlEvents), false);

                    try
                    {
                        util.SendMail("allen.worldlibrary@gmail.com;jason@worldlibrary.org;jdstejskal@gmail.com;support@worldlibrary.org", "support@worldlibrary.org", "support@worldlibrary.org", "", "", "World Public Library Server Error Events", string.Join("\r\n", EmailHtmlEvents), null);
                        /*var fromAddress = new MailAddress("allen.worldlibrary@gmail.com", "From Name");
                        var toAddress = new MailAddress("allen.worldlibrary@gmail.com", "To Name");
                        string fromPassword = "C0c0nut1";
                        string subject = "test";
                        string body = string.Join("\r\n", EmailHtmlEvents);

                        var smtp = new SmtpClient
                        {
                            Host = "smtp.gmail.com",
                            Port = 587,
                            EnableSsl = true,
                            DeliveryMethod = SmtpDeliveryMethod.Network,
                            Credentials = new NetworkCredential(fromAddress.Address, fromPassword),
                            Timeout = 20000
                        };
                        using (var message = new MailMessage(fromAddress, toAddress)
                        {
                            Subject = subject,
                            Body = body
                        })
                        {
                            message.IsBodyHtml = true;
                            smtp.Send(message);
                        }*/
                        Console.WriteLine("Sending.... EventViewerReport to Email....");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Couldn not send Email, Check your Internet and Email Account....");
                    }
                }
                else
                {
                    Console.WriteLine("ERROR: Could not read on IPSToCheck.txt.");
                }            
            }
            catch (Exception e)
            {
                //  requires elevation of privilege
                Console.WriteLine("Error!" + e.ToString());
            }            
        }


        static bool ValidateIPAddress(string ipString)
        {
            bool returnValue = false;
            string[] splitValues = ipString.Split('.');

            try
            {
                if (splitValues.Length == 4)

                {
                    if ((Convert.ToInt16(splitValues[0]) >= 0 && Convert.ToInt16(splitValues[0]) <= 255) &&
                    (Convert.ToInt16(splitValues[1]) >= 0 && Convert.ToInt16(splitValues[1]) <= 255) &&
                    (Convert.ToInt16(splitValues[2]) >= 0 && Convert.ToInt16(splitValues[2]) <= 255) &&
                    (Convert.ToInt16(splitValues[3]) >= 0 && Convert.ToInt16(splitValues[3]) <= 255))
                    {
                        returnValue = true;
                    }

                }

            }
            catch
            {

            }

            return returnValue;
        }
    }
}

