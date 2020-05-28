﻿using VentiDemoBOT.Model;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.BotBuilderSamples;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using AdaptiveCards.Templating;
using AdaptiveCards.Rendering;

namespace VentiDemoBOT.Dialogs.Intents
{
    public class RemoveUserDialog : ComponentDialog
    {
        private const string UserInfo = "value-userInfo";

        public RemoveUserDialog() : base(nameof(RemoveUserDialog))
        {
            AddDialog(new TextPrompt(nameof(TextPrompt), new PromptValidator<string>(async (pvc, ct) => true)));

            AddDialog(new WaterfallDialog(nameof(WaterfallDialog), new WaterfallStep[]
            {
                NameStepAsync,
                DomainStepAsync,
                ConfirmStepAsync,
                ProcessStepAsync,
            }));

            InitialDialogId = nameof(WaterfallDialog);
        }

        private static async Task<DialogTurnResult> NameStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {

            stepContext.Values[UserInfo] = new UserProfile();
            Dictionary<string, object> args = (Dictionary<string, object>)stepContext.Options;

            if (args.ContainsKey("DomainName"))
            {
                var userProfile = (UserProfile)stepContext.Values[UserInfo];
                userProfile.DomainName = args["DomainName"].ToString().Replace(" ","");

                return await stepContext.NextAsync();
            } 
            else
            {
                await stepContext.Context.SendActivityAsync(MessageFactory.Text("Please enter the domain name of the user."), cancellationToken);

                var promptOptions = new PromptOptions { Prompt = MessageFactory.Text("EXAMPLE: upn = johndoe2@student.hogent.be | domain name = johndoe2") };

                return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
            }
            
        }

        private static async Task<DialogTurnResult> DomainStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {

            var userProfile = (UserProfile)stepContext.Values[UserInfo];
            if (userProfile.DomainName == null)
            {
                userProfile.DomainName = (string)stepContext.Result;
            }
            
            Dictionary<string, object> args = (Dictionary<string, object>)stepContext.Options;
            Attachment domainCardAttachment = (Attachment)args["domainCardAttachment"];

            return await stepContext.PromptAsync(
               nameof(TextPrompt),
               new PromptOptions
               {
                   Prompt = new Activity
                   {
                       Type = ActivityTypes.Message,
                       Attachments = new List<Attachment>()
                       {
                          domainCardAttachment
                       },
                   },
               },
               cancellationToken: cancellationToken);

        }

        private static async Task<DialogTurnResult> ConfirmStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var rawValues = (JObject)stepContext.Context.Activity.Value;
            var userProfile = (UserProfile)stepContext.Values[UserInfo];
            userProfile.Domain = (string)rawValues.GetValue("Domain");
            userProfile.UserPrincipalName = userProfile.DomainName + userProfile.Domain;

            Dictionary<string, object> args = (Dictionary<string, object>)stepContext.Options;
            Attachment cardAttachment = (Attachment)args["cardAttachment"];

            var fileRead = cardAttachment.Content.ToString();

            string classData = JsonConvert.SerializeObject(userProfile, Formatting.Indented);

            AdaptiveTransformer transformer = new AdaptiveTransformer();
            cardAttachment.Content = JsonConvert.DeserializeObject(transformer.Transform(fileRead, classData));

            return await stepContext.PromptAsync(
               nameof(TextPrompt),
               new PromptOptions
               {
                   Prompt = new Activity
                   {
                       Type = ActivityTypes.Message,
                       Attachments = new List<Attachment>()
                       {
                          cardAttachment
                       },
                   },
               },
               cancellationToken: cancellationToken);

        }

        private async Task<DialogTurnResult> ProcessStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {

            Dictionary<string, object> args = (Dictionary<string, object>)stepContext.Options;
            TokenResponse tokenResponse = (TokenResponse)args["TokenResponse"];
            var userProfile = (UserProfile)stepContext.Values[UserInfo];

            var rawValues = (JObject)stepContext.Context.Activity.Value;

            int selectValue = (int)rawValues.GetValue("SingleSelectVal");

            if (selectValue == 2)
            {
                await OAuthHelpers.RemoveUserAsync(stepContext.Context, tokenResponse, userProfile.UserPrincipalName);
            }
            else
            {
                await stepContext.Context.SendActivityAsync(MessageFactory.Text($"{userProfile.DomainName} is not getting deleted"));
            }

            return await stepContext.EndDialogAsync(null, cancellationToken);
        }


    }
}
