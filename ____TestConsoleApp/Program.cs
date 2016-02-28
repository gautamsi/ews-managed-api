using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Autodiscover;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            var ex = new ExchangeService(ExchangeVersion.Exchange2013);
            ex.Url = new Uri("https://outlook.office365.com/Ews/Exchange.asmx");
            ex.Credentials = new NetworkCredential("grouptest@mysupport.in", "group@P@ssw0rd");
            //ex.Credentials = new NetworkCredential("gautamsi@microsoft.com", "");

            var iitems = ex.FindItems(WellKnownFolderName.Inbox, new ItemView(1));
            foreach (var item in iitems)
            {
                item.Load();
                if(item.Attachments.Count > 0)
                {
                    item.Attachments[0].Load();
                }
            }
            var xxxx = iitems;
            return;

            EmailMessage msg = new EmailMessage(ex);
            msg.Subject = "Message with Attachments";
            msg.Body = "This message contains four file attachments.";
            msg.ToRecipients.Add("grouptest@mysupport.in");

            msg.Attachments.AddFileAttachment(@"D:\dr-test\textFile.txt");
            msg.SendAndSaveCopy();
            return;








            var PR_TRANSPORT_MESSAGE_HEADERS = new ExtendedPropertyDefinition(0x007D, MapiPropertyType.String);
            var psPropSet = new PropertySet(BasePropertySet.IdOnly, PR_TRANSPORT_MESSAGE_HEADERS);

            var firs = ex.FindItems(WellKnownFolderName.Inbox, new ItemView(1));

            foreach (var itItem in firs.Items)
            {
                itItem.Load(psPropSet);
                Object valHeaders;
                if (itItem.TryGetProperty(PR_TRANSPORT_MESSAGE_HEADERS, out valHeaders))
                {
                    var vv = valHeaders;
                }

                var xv = itItem;
            }


            return;
            var groups = ex.ExpandGroup("blue_ash_email_migration@Singhs.pro");
            var d = ex.ResolveName(groups.Members[0].Address);
            return;
            var pass = ex.GetPasswordExpirationDate("gstest2@singhs.pro");

            return;

            var autod = new AutodiscoverService();
            autod.Credentials = ex.Credentials;
            autod.RedirectionUrlValidationCallback = (uri) =>{ return true; };
            autod.GetUserSettings("gstest@singhspro.onmicrosoft.com", new UserSettingName[] {UserSettingName.ActiveDirectoryServer,UserSettingName.AutoDiscoverSMTPAddress, UserSettingName.CasVersion, UserSettingName.ExternalEcpPhotoUrl, UserSettingName.ExternalEwsUrl,
            UserSettingName.ExternalOABUrl, UserSettingName.MailboxDN, UserSettingName.MailboxVersion, UserSettingName.MobileMailboxPolicy, UserSettingName.RedirectUrl, UserSettingName.UserDisplayName, UserSettingName.UserDN, UserSettingName.UserMSOnline});
            return;

            // Create a list of attendees.
            List <AttendeeInfo> attendees = new List<AttendeeInfo>();

            attendees.Add(new AttendeeInfo()
            {
                SmtpAddress = "gstest@singhspro.onmicrosoft.com",
                AttendeeType = MeetingAttendeeType.Organizer
            });

            attendees.Add(new AttendeeInfo()
            {
                SmtpAddress = "gs@singhspro.onmicrosoft.com",
                AttendeeType = MeetingAttendeeType.Required
            });

            // Specify suggested meeting time options.
            AvailabilityOptions myOptions = new AvailabilityOptions();
            myOptions.MeetingDuration = 60;
            myOptions.MaximumNonWorkHoursSuggestionsPerDay = 0;
            myOptions.GoodSuggestionThreshold = 49;
            myOptions.MinimumSuggestionQuality = SuggestionQuality.Good;
            myOptions.DetailedSuggestionsWindow = new TimeWindow(DateTime.Now.AddDays(4), DateTime.Now.AddDays(5));

            // Return a set of suggested meeting times.
            GetUserAvailabilityResults results = ex.GetUserAvailability(attendees,
                                                                         new TimeWindow(DateTime.Now, DateTime.Now.AddDays(2)),
                                                                             AvailabilityData.Suggestions,
                                                                             myOptions);
            // Display available meeting times.
            Console.WriteLine("Availability for {0} and {1}", attendees[0].SmtpAddress, attendees[1].SmtpAddress);
            Console.WriteLine();

            foreach (Suggestion suggestion in results.Suggestions)
            {
                Console.WriteLine(suggestion.Date);
                Console.WriteLine();
                foreach (TimeSuggestion timeSuggestion in suggestion.TimeSuggestions)
                {
                    Console.WriteLine("Suggested meeting time:" + timeSuggestion.MeetingTime);
                    Console.WriteLine();
                }
            }



            return;

            var contact = ex.ResolveName("gstest", ResolveNameSearchLocation.DirectoryOnly, true, PropertySet.FirstClassProperties);
            //var contact = ex.ResolveName("gautamsi", ResolveNameSearchLocation.DirectoryOnly, true, PropertySet.FirstClassProperties);
            var v = contact;

            return;

            var items = ex.FindItems(WellKnownFolderName.SentItems, new ItemView(3));

            var s = ex.BindToItems(new List<ItemId> { items.Items[0].Id }, PropertySet.FirstClassProperties);
            s[0].Item.Load();
            var t = s[0];
            //ex.RenderingMethod = ExchangeService.RenderingMode.JSON;
            //var f = Folder.Bind(ex, new FolderId(WellKnownFolderName.Calendar));
            var x = "";
        }

    }
}
