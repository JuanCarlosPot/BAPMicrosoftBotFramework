// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
using AdaptiveCards;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Azure.CognitiveServices.Language.LUIS.Runtime;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Linq;
using VentiDemoBOT.Dialogs.Intents;
using Microsoft.Extensions.Options;
using System.Collections.Generic;
using VentiDemoBOT.Model;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace Microsoft.BotBuilderSamples
{
    public class MainDialog : LogoutDialog
    {
        protected readonly ILogger _logger;
        public string LuisModelId { get; }
        public string LuisModelKey { get; }
        public string LuisEndpoint { get; }

        private static readonly string[] _cards =
        {
            Path.Combine(".", "Resources", "CreateUserCard.json"),
            Path.Combine(".", "Resources", "GetUserCard.json"),
            Path.Combine(".", "Resources", "RemoveUserCard.json"),
            Path.Combine(".", "Resources", "DomainUserCard.json")
        };

        public MainDialog(IConfiguration configuration, ILogger<MainDialog> logger)
            : base(nameof(MainDialog), configuration["ConnectionName"])
        {
            LuisModelId = configuration.GetValue<string>("LuisModelId");
            LuisModelKey = configuration.GetValue<string>("LuisModelKey");
            LuisEndpoint = configuration.GetValue<string>("LuisEndpoint");
            _logger = logger;

            AddDialog(new OAuthPrompt(
                nameof(OAuthPrompt),
                new OAuthPromptSettings
                {
                    ConnectionName = ConnectionName,
                    Text = "Please login",
                    Title = "Login",
                    Timeout = 300000, // User has 5 minutes to login
                }));

            AddDialog(new TextPrompt(nameof(TextPrompt)));
            AddDialog(new ChoicePrompt(nameof(ChoicePrompt)));

            AddDialog(new CreateUserDialog());
            AddDialog(new RemoveUserDialog());
            AddDialog(new DomainNameDialog());

            AddDialog(new WaterfallDialog(nameof(WaterfallDialog), new WaterfallStep[]
            {
                PromptStepAsync,
                LoginStepAsync,
                CommandStepAsync,
                ProcessStepAsync,
                FinalStepAsync
            }));

            // The initial child Dialog to run.
            InitialDialogId = nameof(WaterfallDialog);
        }

        private async Task<DialogTurnResult> PromptStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
        }

        private async Task<DialogTurnResult> LoginStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Get the token from the previous step. Note that we could also have gotten the
            // token directly from the prompt itself. There is an example of this in the next method.
            var tokenResponse = (TokenResponse)stepContext.Result;
            if (tokenResponse != null)
            {
                return await stepContext.PromptAsync(nameof(TextPrompt), new PromptOptions { Prompt = MessageFactory.Text("What would you like me to do?") }, cancellationToken);
            }

            await stepContext.Context.SendActivityAsync(MessageFactory.Text("Login was not successful please try again."), cancellationToken);
            return await stepContext.EndDialogAsync();
        }

        private async Task<DialogTurnResult> CommandStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            stepContext.Values["command"] = stepContext.Result;

            // Call the prompt again because we need the token. The reasons for this are:
            // 1. If the user is already logged in we do not need to store the token locally in the bot and worry
            // about refreshing it. We can always just call the prompt again to get the token.
            // 2. We never know how long it will take a user to respond. By the time the
            // user responds the token may have expired. The user would then be prompted to login again.
            //
            // There is no reason to store the token locally in the bot because we can always just call
            // the OAuth prompt to get the token or get a new token if needed.
            return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
        }

        private async Task<DialogTurnResult> ProcessStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            if (stepContext.Result != null)
            {
                // We do not need to store the token in the bot. When we need the token we can
                // send another prompt. If the token is valid the user will not need to log back in.
                // The token will be available in the Result property of the task.
                var tokenResponse = stepContext.Result as TokenResponse;


                // If we have the token use the user is authenticated so we may use it to make API calls.
                if (tokenResponse?.Token != null)
                {
                    //Options dictionary to give with each dialog (storing entities, card attachements, tokenresponse)
                    var args = new Dictionary<string, object> { { "TokenResponse", tokenResponse } };

                    //LUIS insert user utterance and get intent and entities back
                    var utterance = (string)stepContext.Values["command"];
                    var cli = new LUISRuntimeClient(new ApiKeyServiceClientCredentials(LuisModelKey)) { BaseUri = new Uri(LuisEndpoint) };
                    var prediction = await cli.Prediction.ResolveWithHttpMessagesAsync(LuisModelId, utterance);

                    var topScoringIntent = prediction.Body.TopScoringIntent.Intent;
                    var entities = prediction.Body.Entities;

                    //Add luis entities to args
                    foreach (var e in entities)
                    {
                        args.Add(e.Type, e.Entity);
                    }

                    //Add an adaptive card to args based on the intent
                    string intent = Regex.Replace(topScoringIntent, " ", "");
                    string cardFilePath = _cards.Where(c => Regex.IsMatch(c, intent)).FirstOrDefault();

                    if (!string.IsNullOrEmpty(cardFilePath))
                    {
                        Attachment cardAttachment = CreateAdaptiveCardAttachment(cardFilePath);
                        args.Add(nameof(cardAttachment), cardAttachment);
                    }

                    //Adds a general domainCard for prompting the user with a domain choiced (student.hogent.be of hogent.be of ...)
                    Attachment domainCardAttachment
                            = CreateAdaptiveCardAttachment(
                                _cards
                                .Where(c => Regex.IsMatch(c, "DomainUser"))
                                .FirstOrDefault());

                    args.Add(nameof(domainCardAttachment), domainCardAttachment);


                    //If chain that determines based on the intent what dialog to push onto the stack
                    if (topScoringIntent == "Me")
                    {
                        await OAuthHelpers.ListMeAsync(stepContext.Context, tokenResponse);

                    }
                    else if (topScoringIntent == "Send mail")
                    {

                        await stepContext.Context.SendActivityAsync(MessageFactory.Text($"The function isn't enabled"), cancellationToken);

                    }
                    else if (topScoringIntent == "Recent")
                    {
                        await OAuthHelpers.ListRecentMailAsync(stepContext.Context, tokenResponse);

                    }
                    else if (topScoringIntent == "Create User")
                    {

                        return await stepContext.BeginDialogAsync(nameof(CreateUserDialog), args, cancellationToken);

                    }
                    else if (topScoringIntent == "Remove User")
                    {

                        return await stepContext.BeginDialogAsync(nameof(RemoveUserDialog), args, cancellationToken);

                    }
                    else if (topScoringIntent == "Check Licenses")
                    {

                        await OAuthHelpers.CheckLicensesAsync(stepContext.Context, tokenResponse);

                    }
                    else if (topScoringIntent == "Get User" || topScoringIntent == "Disable User" || topScoringIntent == "Enable User")
                    {

                        args.Add(nameof(topScoringIntent), topScoringIntent);
                        return await stepContext.BeginDialogAsync(nameof(DomainNameDialog), args, cancellationToken);

                    }
                    else
                    {
                        await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Sorry, I don't know what you mean"), cancellationToken);
                        await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Currently i have the ability to create/remove/delete/disable/enable a user, check your tenant licenses and get your recent mail."), cancellationToken);

                    }
                }

            }
            else
            {
                await stepContext.Context.SendActivityAsync(MessageFactory.Text("We couldn't log you in. Please try again later."), cancellationToken);
            }

            return await stepContext.NextAsync(cancellationToken: cancellationToken);
        }

        private async Task<DialogTurnResult> FinalStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            await stepContext.Context.SendActivityAsync(MessageFactory.Text("Type anything to give me another command."), cancellationToken);

            return await stepContext.EndDialogAsync(null, cancellationToken);
        }

        private static Attachment CreateAdaptiveCardAttachment(string filePath)
        {
            var adaptiveCardJson = System.IO.File.ReadAllText(filePath);
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = "application/vnd.microsoft.card.adaptive",
                Content = JsonConvert.DeserializeObject(adaptiveCardJson),
            };
            return adaptiveCardAttachment;

        }
    }
}
