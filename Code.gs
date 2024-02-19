function sendEmailOnEdit(e) {
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  if (sheetName == "deployment email demo") {
    var cell = e.range;
    var row = cell.getRow();
    var statusColumn = 19; 
    var emailColumn = 20; 
    var teamNameColumn = 21;
    var developerNameColumn = 22;
    var storyColumn = 4;
    var envNameColumn = 8;
    var addtionalInfoColumn = 23;

    if (cell.getColumn() == statusColumn && cell.getValue() == "completed") {
      var emailAddress = sheet.getRange(row, emailColumn).getValue();
      var userStoryDetails = sheet.getRange(row, storyColumn).getValue();
      var teamName = sheet.getRange(row, teamNameColumn).getValue();
      var envName = sheet.getRange(row, envNameColumn).getValue();
      var developerName = sheet.getRange(row, developerNameColumn).getValue();
      var additionalInfo = sheet.getRange(row, addtionalInfoColumn).getValue();

      const currentDate = new Date();

      // Get the date, hours (in 12-hour format), and minutes separately
      const date = currentDate.toLocaleDateString('en-US', { year: 'numeric', month: '2-digit', day: '2-digit' });
      const hours = currentDate.getHours() % 12 || 12; // Convert to 12-hour format and ensure value is 12 if 0
      const minutes = currentDate.getMinutes().toString().padStart(2, '0');
      const ampm = (hours >= 12) ? 'PM' : 'AM';

      // Create the desired output format
      var formattedDateTime = `${date} at ${hours}:${minutes}${ampm}`;



      //------- email--------------

      if (emailAddress) {
        var subject = teamName.toUpperCase() + " Deployment Completed"; // Customize the subject
        var body = "Hi " + teamName + " Team,\n" + "The deployment in row " + row + " has been successfully completed.\n" + "User Story: " + userStoryDetails + "\t\t\t\t  Environment: " + envName + "\n" + "Deployed at: " + formattedDateTime + "\tRequested by: " + developerName + "\n"; // Customize the body
        MailApp.sendEmail(emailAddress, subject, body);
      }

      //------- chat--------------

      function sendMessage(userStoryDetails, teamName, envName, formattedDateTime, developerName, additionalInfo) {
        // Card content structure
        const message = {
          "cards": [
            {
              "header": {
                "title": "Deployment Notification",
                "subtitle": "A new deployment has been made!",
                "imageUrl": "https://cdn-icons-png.flaticon.com/512/3715/3715206.png",
                "imageAltText": "Avatar for Sasha"
              },
              "sections": [
                {
                  "widgets": [
                    {
                      "textParagraph": {
                        "text": "<font color=\"#1976D2\"><b>User Story:</b></font> " + userStoryDetails
                      }
                    },
                    {
                      "textParagraph": {
                        "text": "<font color=\"#388E3C\"><b>Team Name:</b></font> " + teamName
                      }
                    },
                    {
                      "textParagraph": {
                        "text": "<font color=\"#FBC02D\"><b>Environment:</b></font> " + envName
                      }
                    },
                    {
                      "keyValue": {
                        "topLabel": "Deployed at",
                        "content": formattedDateTime
                      }
                    },
                    {
                      "keyValue": {
                        "topLabel": "Requested by",
                        "content": developerName
                      }
                    },
                    {
                      "textParagraph": {
                        "text": "<font color=\"#757575\"><b>Additional Info For QA Team:</b></font>\n" + additionalInfo
                      }
                    }
                  ]
                }
              ]
            }
          ]
        };

        // Replace URL with your webhook URL
        const webhookUrl = "https://chat.googleapis.com/v1/spaces/xxxxxxxxxx/messages?key=xxxxxxxxxxxxxx-xxxxxxxxxxxxx&token=xxxxxxxxxxxxx-xxxxxxxx"

        // Send the card using UrlFetchApp
        const options = {
          "method": "post",
          "payload": JSON.stringify(message),
          "contentType": "application/json",
          "muteHttpExceptions": true
        };

        try {
          // Send the request using UrlFetchApp
          const response = UrlFetchApp.fetch(webhookUrl, options);
          Logger.log("Message sent to webhook: " + response.getContentText());
        } catch (error) {
          Logger.log("Error sending message: " + error);
        }
      }
      // Call the function with your data
      sendMessage(userStoryDetails, teamName, envName, formattedDateTime, developerName, additionalInfo);
    }
  }
}

