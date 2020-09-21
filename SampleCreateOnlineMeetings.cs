using System;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Graph.Auth;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Threading;



/*
 *1. Register with Microsoft 365 Developer Program for development environment. Add Sample Data packs
 *2. Create App Registration in Azure portal. Copy App ID, and create Client Secret. 
 *3. Add Graph API permissions depending on API. Also add User.Read, Directory.Read, and then complete Admin Consent
 *  
 * 
 * 
 */


namespace Graph_CreateEventOnline
{
    class SampleCreateOnlineMeetings
    {
        static async System.Threading.Tasks.Task Main(string[] args)
        {
			//Setup Auth
            string clientId = "fd0ae99b-0717-4985-90a2-829f473c9c36";
            string tenantID = "c73250cd-044c-4f51-beb8-020d87ee8daa";
            string clientSecret = "SfS2_O_jBDvPM-a526EwG_0rn-n5qJ-xSc";

            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantID)
                .WithClientSecret(clientSecret)
                .Build();

            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);
            GraphServiceClient graphClient = new GraphServiceClient(authProvider);


            try
            {
				var @event = new Event
				{
					Subject = "Meeting with Customer, Customer Name-Omer Amin",
					Body = new ItemBody
					{
						ContentType = BodyType.Html,
						Content = "Add Content here"
					},
					Start = new DateTimeTimeZone
					{
						DateTime = "2020-09-20T12:00:00",
						TimeZone = "Pacific Standard Time"
					},
					End = new DateTimeTimeZone
					{
						DateTime = "2020-09-20T14:00:00",
						TimeZone = "Pacific Standard Time"
					},
					Location = new Location
					{
						DisplayName = "Online"
					},
					Attendees = new List<Attendee>()
				{
					new Attendee
					{
						EmailAddress = new EmailAddress
						{
							Address = "omamin@microsoft.com",
							Name = "Omer Amin"
						},
						Type = AttendeeType.Required
					}
				},
					IsOnlineMeeting = true,
					OnlineMeetingProvider = OnlineMeetingProviderType.TeamsForBusiness

				};

				Task<Event> taskEvent = graphClient.Users["AdeleV@barneyteste5.onmicrosoft.com"].Calendar.Events
					.Request()
					.AddAsync(@event);

				Console.WriteLine("Meeting ID - " + taskEvent.Result.Id);
				Console.WriteLine("WebLink - " + taskEvent.Result.OnlineMeeting.JoinUrl);




				//Sleep for 5 seconds, and then delete the meeting using the meeting ID.
				Console.WriteLine("\nSleep for 5 seconds");
				Thread.Sleep(5000);

				Console.WriteLine("Deleting meeting");
				await graphClient.Users["AdeleV@barneyteste5.onmicrosoft.com"].Calendar.Events[taskEvent.Result.Id]
					.Request()
					.DeleteAsync();

			}
			catch (ServiceException e)
            {
				Console.WriteLine(e.Message);
				Console.WriteLine(e.Error.Code);
				Console.WriteLine(e.StackTrace);
				throw;
            }


			//Specify meeting parameters, and then Add to user's calendar. This will return the Teams meeting link in OnlineMeeting.JoinUrl property.



			Console.WriteLine("Meeting Deleted");

		}


    }
}
