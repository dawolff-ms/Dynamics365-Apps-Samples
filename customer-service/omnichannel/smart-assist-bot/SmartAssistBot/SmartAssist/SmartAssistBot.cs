// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
using AdaptiveCards;
using CoreBot.Models;
using CoreBot.SmartAssist.Operations;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Microsoft.BotBuilderSamples.Bots
{
    public class ConversationData
    {
        public bool IsEscalated { get; set; }

    }

    /// <summary>
    /// Smartassist bot to demonstrate knowledge article suggestions and appointment intent detection
    /// </summary>
    public class SmartAssistBot : ActivityHandler
    {
        // Data image URIs used for convenience. You can use full URLs for images
        private const string AppointmentIcon = "data:image/png;base64,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";
        private const string CreateIcon = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAA4AAAANCAYAAACZ3F9/AAAACXBIWXMAAAsTAAALEwEAmpwYAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAE8SURBVHgBnVLBUQJBEOxd1z8BALWgAWAEagSQAfLw5CdEIBkIP+UeEoIZyGUAAchtnVT55cXDu5txDrkqQKHUeexu7Uz3dO8OrBc18I/QWul6xYtC682bfwGq6nV0xazqSqkFg63sE9LxwD1U3MGOdJSOoVCY+cUWx6YFpqmm45fqTfRk26Hd2zFbqt48pI/4zI0qizyxUqJVU1NWxYPXYfl5q+MaPoYxtc3EzC+PwsfSZWrQY/7+Dipnl+M9xOc+aWAuZJYYmDCha7766nMibrhhMditt52woJemI6RNqekjSQaZJbNmqyFOppuA0/bbhTDf0lJyjD7F229gbPvdcpogv1wBGHdpKnxAz/nl4Cfl6kQmhxXqIDUROR35jjFBjZxfCnAgDMvkQPwRw2FHzkEgiAKK0+5vAXl8ArWLn19rFeLfAAAAAElFTkSuQmCC";

        protected readonly BotState ConversationState;
        protected readonly BotState UserState;
        protected readonly ILogger Logger;
        public const string OmnichannelFooBar = "omnichannelfoobar";

        public SmartAssistBot(ConversationState conversationState, UserState userState, IDynamicsDataAccessLayer dynamicsDataAccessLayer, KBSearchOperation kBSearchOperation, AppointmentDetectionOperation appointmentDetectionOperation)
        {
            ConversationState = conversationState;
            UserState = userState;
        }

        private async Task setConversationData(ITurnContext turnContext)
        {
            var conversationStateAccessors = ConversationState.CreateProperty<ConversationData>(nameof(ConversationData));
            var conversationData = await conversationStateAccessors.GetAsync(turnContext, () => new ConversationData());

            conversationData.IsEscalated = true;
            await ConversationState.SaveChangesAsync(turnContext);
            await UserState.SaveChangesAsync(turnContext);
        }

        public override async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default(CancellationToken))
        {
            var conversationStateAccessors = ConversationState.CreateProperty<ConversationData>(nameof(ConversationData));
            var conversationData = await conversationStateAccessors.GetAsync(turnContext, () => new ConversationData());

            var activity = turnContext.Activity;
            var actList = new List<IActivity>();

            if (activity.Type == ActivityTypes.Message &&
                activity.Text ==  OmnichannelFooBar &&
                conversationData.IsEscalated == false)
            {
                await setConversationData(turnContext);
                return;
            }

            if (activity.Type == ActivityTypes.ConversationUpdate && conversationData.IsEscalated == false)
            {
                var result = activity.MembersAdded;
                foreach (ChannelAccount account in result)
                {
                    if (account.AadObjectId != null)
                    {
                        await setConversationData(turnContext);
                        return;
                    }
                }
            }
            if (activity.Type == ActivityTypes.Message && (conversationData.IsEscalated == true))
            {
                await base.OnTurnAsync(turnContext, cancellationToken);
            }

            // Save any state changes that might have occured during the turn.
            await ConversationState.SaveChangesAsync(turnContext, false, cancellationToken);
            await UserState.SaveChangesAsync(turnContext, false, cancellationToken);
        }

        protected AdaptiveCard ConstructionAppointmentCard()
        {
            // Constructs the adaptive card similar to /SmartAssist/Cards/CreateAppointment.json
            var appointmentAdaptiveCard = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>()
                {
                    new AdaptiveContainer() {
                        Items = new List<AdaptiveElement>() {
                            new AdaptiveColumnSet() {
                                Columns = new List<AdaptiveColumn>() {
                                    new AdaptiveColumn() {
                                        Width = "auto",
                                        Items = new List<AdaptiveElement>() {
                                            new AdaptiveImage() {
                                                PixelWidth = 16,
                                                PixelHeight = 16,
                                                Url = new System.Uri(AppointmentIcon),
                                                AltText = "Appointment"
                                            }
                                        }
                                    },
                                    new AdaptiveColumn() {
                                        Width = "stretch",
                                        Items = new List<AdaptiveElement>() {
                                            new AdaptiveTextBlock() {
                                                Text = "Appointment"
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    },
                    new AdaptiveContainer() {
                        Items = new List<AdaptiveElement>() {
                            new AdaptiveTextBlock() {
                                Text = "Phone call with customer",
                                Weight = AdaptiveTextWeight.Bolder
                            }
                        }
                    },
                    new AdaptiveTextBlock() {
                        Text = "Setup a phone call with the customer. The appointment fields will be auto populated based on context",
                        Wrap = true
                    }
                },
                Actions = new List<AdaptiveAction>()
                {
                    new AdaptiveSubmitAction() {
                        IconUrl = CreateIcon,
                        Title = "Create appointment",
                        Data = new CustomActionData() {
                            // This custom action should already be available and accessible to Smart assist control
                            CustomAction = "CreateAppointment",

                            // Always pass an empty dictionary even if there are no parameters
                            CustomParameters = new Dictionary<string, string>()
                        }
                    }
                }
            };
            return appointmentAdaptiveCard;
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            // Always create an appointment card.
            var cards = new List<IMessageActivity>();
            var appointmentAdaptiveCard = ConstructionAppointmentCard();
            var attachment = new Attachment()
            {
                ContentType = "application/vnd.microsoft.card.adaptive",
                Content = appointmentAdaptiveCard,
            };
            var card = MessageFactory.Attachment(attachment);
            cards.Add(card);

            // IMPORTANT!
            // This tag MUST be present in all responses going from the bot
            // Any response without this tag will cause unintended UX errors
            cards.ForEach((card) =>
            {
                Dictionary<string, object> channelinfo = new Dictionary<string, object>
                        {
                            { "tags", "smartbot" }
                        };
                card.ChannelData = channelinfo;
            });

            // Send the card(s) as response
            await turnContext.SendActivitiesAsync(cards.ToArray(), cancellationToken);
        }
    }
}
