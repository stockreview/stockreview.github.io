/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

/// <reference path="/Scripts/jquery.fabric.js" />

(function () {
    'use strict';

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
	//	Office.onReady((info) => {
        $(document).ready(function () {

            detectActionsForMe(Office.context.mailbox.item);
			//event handler for item change event (i.e. new message selected)
			Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

			UpdateTaskPaneUI(Office.context.mailbox.item);
        });
		
    };

    function detectActionsForMe(item) {
        //var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
		$('#to-result').text('no');
		
		// add something to the infobar
		Office.context.mailbox.item.notificationMessages.addAsync("information", {
		type: "informationalMessage",
		message : "My custom message.",
		icon : "iconid",
		persistent: true
		});
		
		
		
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            var myInfo = Office.context.mailbox.userProfile;

			var nameParts = myInfo.displayName.split(' ');

			// Make an assumption that the displayName is in the form 'firstName lastName'
			var myFirstName = nameParts[0];

			// Check whether I was cc'd on this email
			for (var i = 0; i < item.cc.length; i++) {
				if (item.cc[i].emailAddress === myInfo.emailAddress) {
					$('#cc-result').text('yeso');
					break;
				}
			}

			// Check whether I am on the To line
			for (var i = 0; i < item.to.length; i++) {
				if (item.to[i].emailAddress === myInfo.emailAddress) {
					$('#to-result').text('yeso');
					break;
				}
			}
			$('#to-result').text("hi" + item.from.emailAddress);
			// Check whether external sender
			//if (item.from.emailAddress.substr(item.from.emailAddress.indexOf("@")) === myInfo.emailAddress.substr(myInfo.emailAddress.indexOf("@"))) {
			//		$('#to-result').text('internal');
			//	}
			//	else {
			//		$('#to-result').text('external');
			//	}

			if (((item.from.displayName.indexOf("@") > 0 ) && (item.from.displayName != item.from.emailAddress)) || ((item.from.displayName.indexOf(".") > 0 ) &&  (item.from.displayName.indexOf(" ") < 0 ) &&  (item.from.displayName.indexOf("@") < 0 )))  {
					$('#cc-result').text('failed' + item.from.displayName.indexOf(" "));
				}
				else {
					$('#cc-result').text('passed');
				}
			// We need to determine if body.getAsync() is defined. We require this method in order to
			// retrieve the body test for parsing. This method was added in v1.3 of the API
			// and may not be available on every Outlook client.
			//
			// For more information, please see Understanding API Requirement Sets at
			// https://dev.outlook.com/reference/add-ins/tutorial-api-requirement-sets.html
			if (Office.context.mailbox.item.body.getAsync !== undefined) {
				// Check whether I am mentioned in the body of the email by name
				// In this sample we scan the email body as plain text. You can also
				// set the coercionType on the getAsync() method to retrieve the body as HTML.
				// For an example of retrieving the body as HTML and parsing the result,
				// see https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.js
				Office.context.mailbox.item.body.getAsync('text', function (asyncResult) {
					var bodyText = asyncResult.value;

					// Create regular expression to find all matches of the first name
					// i => ignore case
					// g => global match, i.e., doesn't stop after first match
					var regex = new RegExp(myFirstName, 'gi');
					var matchingArray = new Array();
					while (regex.exec(bodyText)) {
						matchingArray.push(regex.lastIndex);
					}

					var result = 'Scan Completeo.';

					if (matchingArray.length > 0) {
						showNotification('Scan Complete', 'It looks like you are mentioned by name in the body of this email');
					}
					else {
						showNotification('Scan Complete', 'It looks like you are not mentioned by name in the body of this email');
					}

				});
			}
			else { // Method not available
				showNotification('Warning', 'The body.getAsync() method is not available in this version of Outlook. Body parsing was skipped');
			}
         
        }
    }
	
	function showNotification(header, content) {

        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);

        var element = document.querySelector('.ms-MessageBanner');
 //       var messageBanner = new fabric.MessageBanner(element);
 //       messageBanner.showBanner();
    }
	
function UpdateTaskPaneUI(item)
{
  // Assuming that item is always a read item (instead of a compose item).
  if (item != null) detectActionsForMe(item);
}
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
    
})();

// *********************************************************
//
// Outlook-Add-in-ScanForMe, https://github.com/OfficeDev/Outlook-Add-in-ScanForMe
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************
