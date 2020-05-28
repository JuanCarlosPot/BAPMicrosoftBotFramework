using VentiDemoBOT.Model;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.BotBuilderSamples;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace VentiDemoBOT.Dialogs.Intents
{
    public class CreateUserDialog : ComponentDialog
    {

        public CreateUserDialog() : base(nameof(CreateUserDialog))
        {
            AddDialog(new TextPrompt(nameof(TextPrompt), new PromptValidator<string>(async (pvc, ct) => true)));

            AddDialog(new WaterfallDialog(nameof(WaterfallDialog), new WaterfallStep[]
            {
                GetInformationStepAsync,
                ProcessStepAsync
            }));

            InitialDialogId = nameof(WaterfallDialog);
        }

        private static async Task<DialogTurnResult> GetInformationStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {

            Dictionary<string, object> args = (Dictionary<string, object>)stepContext.Options;
            Attachment cardAttachment = (Attachment)args["cardAttachment"];

            var reply = MessageFactory.Attachment(cardAttachment);

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
            var userProfile = new UserProfile();
            

            //Contains the inputed variables
            var rawValues = (JObject)stepContext.Context.Activity.Value;
            
            foreach (var val in rawValues)
            {
                if (val.Key != "Domain")
                {
                    userProfile.GetType().GetProperty(val.Key.ToString()).SetValue(userProfile, val.Value.ToString());
                }
                else
                {
                    userProfile.Domain = val.Value.ToString();
                }
            }

            userProfile.DomainName = userProfile.GivenName.Replace(" ","") + userProfile.SurName.Replace(" ", "");
            userProfile.UserPrincipalName = userProfile.DomainName + userProfile.Domain;

            await OAuthHelpers.CreateUserAsync(stepContext.Context, tokenResponse, userProfile);

            return await stepContext.EndDialogAsync(null, cancellationToken);
        }
    }
}
