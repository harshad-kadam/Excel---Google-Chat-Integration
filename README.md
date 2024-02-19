# Automated Notification for Excel

## Description

Automated Notification for Excel leverages Google Apps Script to streamline communication by automatically sending notifications upon specific changes in your Excel sheet. This solution allows you to configure notifications via email and Google Chat groups, ensuring stakeholders stay informed in real-time.

## Key Features

- **Trigger-based**: Notifications are sent when a designated cell is edited, ensuring timely updates.
- **Customizable**: Define email addresses, Google Chat group IDs, and message content to match your needs.
- **Flexible**: Adapt the code to different cell addresses, triggers, and notification methods.
- **Extensible**: Explore other integration possibilities like SMS, Slack, or custom webhooks.

## Prerequisites

- A Google Account
- A Google Sheet

## Setup

1. **Create a Google Sheet**: Design your sheet with the data you want to monitor and a dedicated "Notification Trigger" cell (e.g., A1).
[Excel Notifications sheet](https://docs.google.com/spreadsheets/d/1DmWgvaK1L3158-7zOx85wRlJ-LZ4_Pdes1FPky0-nMk/edit?usp=sharing)

3. **Add Apps Script**:
   - Go to Tools > Script editor.
   - Paste the provided code, replacing placeholders with your values.
   - Save the project (e.g., "Excel Notifications").

4. **Authorize the Script**:
   - Run the script and grant necessary permissions when prompted.

5. **Set Trigger (Optional)**:
   - Go to Resources > Current project's triggers.
   - Create a time-driven trigger to run the script periodically, catching notifications even without direct edits.

## Code (with Placeholders)

```javascript
// Placeholders are indicated with comments (// Replace with ...)
// Update placeholders with actual values before running the script
// Use code with caution.

// Function to run on every edit in the sheet
function onEdit(e) {
  // Replace with actual sheet name and cell address
  const sheetName = "Notification Sheet";
  const cellAddress = "A1";

  // Check if edited sheet and cell match expectations
  if (e.source.getActiveSheet().getName() !== sheetName || e.range.getAddress() !== cellAddress) {
    return;
  }

  // Retrieve data from the sheet
  // Replace cell addresses with actual ones
  const userStoryDetails = sheetName.getRange("B1").getValue();
  const teamName = sheetName.getRange("C1").getValue();
  const envName = sheetName.getRange("D1").getValue();
  const formattedDateTime = sheetName.getRange("E1").getValue(); // Example: 2024-02-21 17:31

  // Placeholders for customization
  const emailAddresses = ["recipient1@example.com", "recipient2@example.com"];
  const googleChatGroupId = "AAAABV0jd7w"; // Replace with your group ID
  const subject = "Notification - Story ${userStoryDetails}";

  // Compose message content
  const messageParts = {
    story: userStoryDetails,
    team: teamName,
    environment: envName,
    deployedAt: formattedDateTime
  };
  const formattedMessage = `
    *Story:* ${messageParts.story}
    *Team:* ${messageParts.team}
    Deployed to *${messageParts.environment}* at *${messageParts.deployedAt}*.
  `;

  // Send email notification (example using GmailApp)
  try {
    GmailApp.sendEmail(emailAddresses, subject, formattedMessage);
    console.log("Email notification sent successfully.");
  } catch (error) {
    console.error("Error sending email:", error);
  }

  // Send Google Chat notification using UrlFetchApp
  const payload = {
    text: formattedMessage
  };
  const options = {
    "method" : "post",
    "payload" : JSON.stringify(payload),
    "headers": {
      "Content-Type": "application/json"
    }
  };

  try {
    UrlFetchApp.fetch(`https://chat.googleapis.com/v1/spaces/${googleChatGroupId}/messages`, options);
    console.log("Google Chat notification sent successfully.");
  } catch (error) {
    console.error("Error sending Google Chat message:", error);
  }
}
```
#### Curl For Reference
in below curl replace URI after importing curl(use URI of google chat group added webhook)

to add webhook [refer it here](https://docs.google.com/document/d/1xgwvl5bO1LJ2QjKZk56vmrU0amO6m6N57GaVNZOvUq0/edit?usp=sharing)

final notification format with provided [JS code](https://docs.google.com/document/d/1HBHAZseNiOLHoR6y542zMOUV0lvhpoBR9u98duTNZCs/edit?usp=sharing)

```c
curl --location 'https://chat.googleapis.com/v1/spaces/xxxxxxxxxx/messages?key=xxxxxxxxxx-xxxxxxxxxxxxxxxxxxx&token=xxxxxxxxxxxxxxxxx-xxxxxxxxx' \
--header 'Content-Type: application/json' \
--data-raw '{
    "cardsV2": [
        {
            "cardId": "unique-card-id",
            "card": {
                "header": {
                    "title": "Sasha",
                    "subtitle": "Software Engineer",
                    "imageUrl": "https://developers.google.com/chat/images/quickstart-app-avatar.png",
                    "imageType": "CIRCLE",
                    "imageAltText": "Avatar for Sasha"
                },
                "sections": [
                    {
                        "header": "Contact Info",
                        "collapsible": true,
                        "uncollapsibleWidgetsCount": 1,
                        "widgets": [
                            {
                                "decoratedText": {
                                    "startIcon": {
                                        "knownIcon": "EMAIL"
                                    },
                                    "text": "sasha@example.com"
                                }
                            },
                            {
                                "decoratedText": {
                                    "startIcon": {
                                        "knownIcon": "PERSON"
                                    },
                                    "text": "<font color=\"#80e27e\">Online</font>"
                                }
                            },
                            {
                                "decoratedText": {
                                    "startIcon": {
                                        "knownIcon": "PHONE"
                                    },
                                    "text": "+1 (555) 555-1234"
                                }
                            },
                            {
                                "buttonList": {
                                    "buttons": [
                                        {
                                            "text": "Share",
                                            "onClick": {
                                                "openLink": {
                                                    "url": "https://example.com/share"
                                                }
                                            }
                                        },
                                        {
                                            "text": "Edit",
                                            "onClick": {
                                                "action": {
                                                    "function": "goToView",
                                                    "parameters": [
                                                        {
                                                            "key": "viewType",
                                                            "value": "EDIT"
                                                        }
                                                    ]
                                                }
                                            }
                                        }
                                    ]
                                }
                            }
                        ]
                    }
                ]
            }
        }
    ]
}'
```
## Customization

- Update placeholders with your data sources, notification preferences, and formatting.
- Add more data points from your sheet to the message content.
- Implement conditional logic to trigger notifications based on specific cell values or changes.

## What can be improved

- **Error Handling**: Implement robust error handling to gracefully manage exceptions and failures during notification delivery.
  
- **Enhanced Trigger Options**: Explore additional trigger options like cell value changes or specific criteria for notifications.
  
- **Integration Expansion**: Extend integration possibilities to include other communication platforms such as SMS, Slack, or custom webhooks.

## Installation

1. Clone the repository:
   git clone git@github.com:harshad-kadam/Excel-Google-Chat-Integration.git

2. Create feature_yourname branch & git checkout to feature branch

3. Open folder with VSCode

4. Add your changes 

5. Git commit & push changes

6. Create PR n share on kadamharshad25@gmail.com

7. I will merge to main code. 

8. You did it. 🏆

9. Join Our channel https://t.me/apigeedeveloper

# 😊Thanks for being here🚀
