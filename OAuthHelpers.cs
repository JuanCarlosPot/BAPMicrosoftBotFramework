// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using VentiDemoBOT.Model;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.Graph;
using Newtonsoft.Json;
using Attachment = Microsoft.Bot.Schema.Attachment;
using AdaptiveCards;
using Newtonsoft.Json.Linq;
using AdaptiveCards.Templating;


namespace Microsoft.BotBuilderSamples
{
    // This class calls the Microsoft Graph API. The following OAuth scopes are used:
    // 'OpenId' 'email' 'Mail.Send.Shared' 'Mail.Read' 'profile' 'User.Read' 'User.ReadBasic.All'
    // for more information about scopes see:
    // https://developer.microsoft.com/en-us/graph/docs/concepts/permissions_reference
    public static class OAuthHelpers
    {

        // Enable the user to send an email via the bot.
        public static async Task SendMailAsync(ITurnContext turnContext, TokenResponse tokenResponse, string recipient)
        {
            if (turnContext == null)
            {
                throw new ArgumentNullException(nameof(turnContext));
            }

            if (tokenResponse == null)
            {
                throw new ArgumentNullException(nameof(tokenResponse));
            }

            var client = new SimpleGraphClient(tokenResponse.Token);
            var me = await client.GetMeAsync();

            await client.SendMailAsync(
                recipient,
                "Message from a bot!",
                $"Hi there! I had this message sent from a bot. - Your friend, {me.DisplayName}");

            await turnContext.SendActivityAsync(
                $"I sent a message to '{recipient}' from your account.");
        }

        // Displays information about the user in the bot.
        public static async Task ListMeAsync(ITurnContext turnContext, TokenResponse tokenResponse)
        {
            if (turnContext == null)
            {
                throw new ArgumentNullException(nameof(turnContext));
            }

            if (tokenResponse == null)
            {
                throw new ArgumentNullException(nameof(tokenResponse));
            }

            // Pull in the data from the Microsoft Graph.
            var client = new SimpleGraphClient(tokenResponse.Token);
            var me = await client.GetMeAsync();

            await turnContext.SendActivityAsync($"You are {me.DisplayName}.");
        }

        // Gets recent mail the user has received within the last hour and displays up
        // to 5 of the emails in the bot.
        public static async Task ListRecentMailAsync(ITurnContext turnContext, TokenResponse tokenResponse)
        {
            if (turnContext == null)
            {
                throw new ArgumentNullException(nameof(turnContext));
            }

            if (tokenResponse == null)
            {
                throw new ArgumentNullException(nameof(tokenResponse));
            }

            var client = new SimpleGraphClient(tokenResponse.Token);
            var messages = await client.GetRecentMailAsync();
            IMessageActivity reply = null;

            if (messages.Any())
            {
                var count = messages.Length;
                if (count > 5)
                {
                    count = 5;
                }

                reply = MessageFactory.Attachment(new List<Attachment>());
                reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

                for (var i = 0; i < count; i++)
                {
                    var mail = messages[i];
                    var card = new HeroCard(
                        mail.Subject,
                        $"{mail.From.EmailAddress.Name} <{mail.From.EmailAddress.Address}>",
                        mail.BodyPreview,
                        new List<CardImage>()
                        {
                            new CardImage(
                                "https://botframeworksamples.blob.core.windows.net/samples/OutlookLogo.jpg",
                                "Outlook Logo"),
                        });
                    reply.Attachments.Add(card.ToAttachment());
                }
            }
            else
            {
                reply = MessageFactory.Text("Unable to find any recent unread mail.");
            }

            await turnContext.SendActivityAsync(reply);

        }

        public static async Task CheckLicensesAsync(ITurnContext turnContext, TokenResponse tokenResponse)
        {
            if (turnContext == null)
            {
                throw new ArgumentNullException(nameof(turnContext));
            }

            if (tokenResponse == null)
            {
                throw new ArgumentNullException(nameof(tokenResponse));
            }

            var client = new SimpleGraphClient(tokenResponse.Token);
            IMessageActivity reply = null;

            //NO RIGHTS AVAILABLE IN DEMO
            //var licenses = await client.GetLicensesAsync();

            //if (licenses.Any())
            //{
            //    var count = licenses.Length;
            //    reply = MessageFactory.Attachment(new List<Attachment>());
            //    reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

            //    for (var i = 0; i < count; i++)
            //    {
            //        var license = licenses[i];
            //        var card = new HeroCard(
            //            license.SkuPartNumber,
            //            license.SkuId.ToString(),
            //            "INFO ABOUT LICENSE (api can't include expiry date)",
            //            new List<CardImage>()
            //            {
            //                new CardImage(
            //                    "https://us.123rf.com/450wm/aquir/aquir2003/aquir200308059/143107865-subscription-stamp-subscription-vintage-gray-label-sign.jpg?ver=6",
            //                    "Outlook Logo"),
            //            });
            //        reply.Attachments.Add(card.ToAttachment());
            //    }
            //}
            //else
            //{
            //    reply = MessageFactory.Text("Unable to find any recent unread mail.");
            //}


            //COMMENT THIS SECTION WHEN IMPLEMENTION ABOVE IS UNCOMMENTEDD
            reply = MessageFactory.Attachment(new List<Attachment>());
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

            for (var i = 0; i < 5; i++)
            {
                //would replace with properties out of var licenses
                var card = new HeroCard(
                    "Enterprise E5 (with Audio Conferencing)", "ENTERPRISEPREMIUM",
                    "INFO ABOUT LICENSE (api can't include expiry date)",
                    new List<CardImage>()
                    {
                            new CardImage(
                                "https://www.infserv.com/images/Artikels/office2.jpg",
                                "Office365 Logo"),
                    });
                reply.Attachments.Add(card.ToAttachment());
            }
            //COMMENT THIS SECTION WHEN IMPLEMENTION ABOVE IS UNCOMMENTEDD

            await turnContext.SendActivityAsync(reply);
            await turnContext.SendActivityAsync(MessageFactory.Text("This is a DEMO, these cards are made from fictional data"));

        }

        public static async Task DisableUserAsync(ITurnContext turnContext, TokenResponse tokenResponse, string upn)
        {
            if (turnContext == null)
            {
                throw new ArgumentNullException(nameof(turnContext));
            }

            if (tokenResponse == null)
            {
                throw new ArgumentNullException(nameof(tokenResponse));
            }

            var client = new SimpleGraphClient(tokenResponse.Token);

            //need to check if the function executed correctly if it were to be implemented
            //await client.DisableUserAsync(upn);

            await turnContext.SendActivityAsync($"The user: {upn} has been DISABLED.");
            await turnContext.SendActivityAsync(MessageFactory.Text("This is a DEMO, no write actions were executed"));
        }

        public static async Task EnableUserAsync(ITurnContext turnContext, TokenResponse tokenResponse, string upn)
        {
            if (turnContext == null)
            {
                throw new ArgumentNullException(nameof(turnContext));
            }

            if (tokenResponse == null)
            {
                throw new ArgumentNullException(nameof(tokenResponse));
            }

            var client = new SimpleGraphClient(tokenResponse.Token);

            //Function returns null NOT IMPLEMENTED
            //need to check if the function executed correctly if it were to be implemented
            //await client.EnableUserAsync(upn);

            await turnContext.SendActivityAsync($"The user: {upn} has been ENABLED.");
            await turnContext.SendActivityAsync(MessageFactory.Text("This is a DEMO, no write actions were executed"));
        }

        public static async Task GetUserAsync(ITurnContext turnContext, TokenResponse tokenResponse, string upn)
        {
            if (turnContext == null)
            {
                throw new ArgumentNullException(nameof(turnContext));
            }

            if (tokenResponse == null)
            {
                throw new ArgumentNullException(nameof(tokenResponse));
            }

            var client = new SimpleGraphClient(tokenResponse.Token);
            var user = await client.GetUserAsync(upn);

            var fileRead = System.IO.File.ReadAllText(@".\Resources\GetUserCard.json");
            var item = (JObject)JsonConvert.DeserializeObject(fileRead);

            string classData = JsonConvert.SerializeObject(user, Formatting.Indented);

            AdaptiveTransformer transformer = new AdaptiveTransformer();
            string cardJson = transformer.Transform(fileRead, classData);

            Attachment attachment = new Attachment();
            attachment.ContentType = "application/vnd.microsoft.card.adaptive";
            attachment.Content = JsonConvert.DeserializeObject(cardJson);

            var attachments = new List<Attachment>();
            var reply = MessageFactory.Attachment(attachments);
            reply.Attachments.Add(attachment);

            await turnContext.SendActivityAsync(reply);
        }

        public static async Task CreateUserAsync(ITurnContext turnContext, TokenResponse tokenResponse, UserProfile userProfile)
        {
            if (turnContext == null)
            {
                throw new ArgumentNullException(nameof(turnContext));
            }

            if (tokenResponse == null)
            {
                throw new ArgumentNullException(nameof(tokenResponse));
            }

            var client = new SimpleGraphClient(tokenResponse.Token);

            User cUser = new User {
                AccountEnabled = true,
                DisplayName = userProfile.GivenName + " " + userProfile.SurName,
                UserPrincipalName = userProfile.UserPrincipalName, 
                GivenName = userProfile.GivenName, 
                Surname = userProfile.SurName, 
                Department = userProfile.Department, 
                JobTitle = userProfile.JobTitle ,
                PasswordProfile = new PasswordProfile
                {
                    ForceChangePasswordNextSignIn = true,
                    Password = userProfile.Password
                }
            };

            //need to check if the function executed correctly if it were to be implemented
            //await client.CreateUserAsync(cUser);

            await turnContext.SendActivityAsync(MessageFactory.Text($"{JsonConvert.SerializeObject(cUser, Formatting.Indented)}"));
            await turnContext.SendActivityAsync(MessageFactory.Text("This is a DEMO, no write actions were executed"));

        }

        public static async Task RemoveUserAsync(ITurnContext turnContext, TokenResponse tokenResponse, string upn)
        {
            if (turnContext == null)
            {
                throw new ArgumentNullException(nameof(turnContext));
            }

            if (tokenResponse == null)
            {
                throw new ArgumentNullException(nameof(tokenResponse));
            }

            var client = new SimpleGraphClient(tokenResponse.Token);

            //need to check if the function executed correctly if it were to be implemented
            //await client.RemoveUserAsync(upn);

            await turnContext.SendActivityAsync($"The user: {upn} has been REMOVED.");
            await turnContext.SendActivityAsync(MessageFactory.Text("This is a DEMO, no write actions were executed"));
        }


    }
}
