using VentiDemoBOT.Model;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.BotBuilderSamples;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace VentiDemoBOT.Dialogs.Intents
{
    public class DomainNameDialog : ComponentDialog
    {
        private const string UserInfo = "value-userInfo";

        public DomainNameDialog() : base(nameof(DomainNameDialog))
        {
            AddDialog(new TextPrompt(nameof(TextPrompt), new PromptValidator<string>(async (pvc, ct) => true)));

            AddDialog(new WaterfallDialog(nameof(WaterfallDialog), new WaterfallStep[]
            {
                NameStepAsync,
                DomainStepAsync,
                ProcessStepAsync
            }));

            InitialDialogId = nameof(WaterfallDialog);
        }

        private static async Task<DialogTurnResult> NameStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {

            stepContext.Values[UserInfo] = new UserProfile();
            Dictionary<string, object> args = (Dictionary<string, object>)stepContext.Options;

            //Checks if the domainname has already been given as a LUIS entity
            if (args.ContainsKey("DomainName"))
            {
                var userProfile = (UserProfile)stepContext.Values[UserInfo];
                userProfile.DomainName = args["DomainName"].ToString().Replace(" ", "");

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

        private async Task<DialogTurnResult> ProcessStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var rawValues = (JObject)stepContext.Context.Activity.Value;
            var userProfile = (UserProfile)stepContext.Values[UserInfo];
            userProfile.Domain = (string)rawValues.GetValue("Domain");
            userProfile.UserPrincipalName = userProfile.DomainName + userProfile.Domain;

            Dictionary<string, object> args = (Dictionary<string, object>)stepContext.Options;

            string topScoringIntent = (string)args["topScoringIntent"];
            TokenResponse tokenResponse = (TokenResponse)args["TokenResponse"];
            

            if (topScoringIntent == "Get User")
            {
                Attachment cardAttachment = (Attachment)args["cardAttachment"];
                await OAuthHelpers.GetUserAsync(stepContext.Context, tokenResponse, userProfile.UserPrincipalName);

            }
            else if (topScoringIntent == "Disable User")
            {

                await OAuthHelpers.DisableUserAsync(stepContext.Context, tokenResponse, userProfile.UserPrincipalName);

            }
            else if (topScoringIntent == "Enable User")
            {

                await OAuthHelpers.EnableUserAsync(stepContext.Context, tokenResponse, userProfile.UserPrincipalName);

            }

            return await stepContext.EndDialogAsync(null, cancellationToken);
        }
    }
}
