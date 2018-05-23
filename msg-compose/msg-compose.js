/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

"use strict";
var client = "unknown";

(function() {
  $(document).ready(function() {
    //  $('#run').click(run);
    
    // Office.context.mailbox.item.notificationMessages.replaceAsync("progress1", {
    //   type: "errorMessage",
    //   message : "Foo Test"
    //   });

    var CheckBoxElements = document.querySelectorAll(".ms-CheckBox");
    for (var i = 0; i < CheckBoxElements.length; i++) {
      new fabric["CheckBox"](CheckBoxElements[i]);
    }

    var ButtonElements = document.querySelectorAll(".ms-Button");
    for (var i = 0; i < ButtonElements.length; i++) {
      new fabric["Button"](ButtonElements[i], function() {
        // Insert Event Here
      });
    }

    // var inBtn = document.getElementById("insert-button");
    // inBtn.onclick = insertSelectedForms;
    $('#insert-button').on('click', insertSelectedForms);
    write("wired up insert button");

    function logResults(json){
      console.log(json);
    }

    //headMsg.textContent = "username to go here";

    var requestUrl = 'http://localhost:50268/api/UserInfo';

    // This works in a browswer
    // $.ajax({
    //   url: requestUrl,
    //   xhrFields: {
    //     withCredentials: true
    //   },
    //   crossDomain: true
    //   //headers: {'Authorization': 'Bearer ' + accessToken}
    // }).done(function(userInfo){
    //   headMsg.textContent = "User: " + userInfo.firstName + " " + userInfo.lastName + " Dept: " + userInfo.departmentName;
    // }).fail(function(jqXHR, textStatus, errorThrown){
    //   headMsg.textContent = "failed textstatus: [" + textStatus + "]  error: [" + errorThrown + "]";
    // });

    

  });

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function(reason) {
    client = "could be office";
  };

  function run() {
    /**
     * Insert your Outlook code here
     */
  }

  function insertSelectedForms(obj) {
    obj.currentTarget.style.backgroundColor = "red";
    alert("here we are");
    write('insert clicked ...');


    // Determine current browser or outlook
    var headMsg = document.getElementById("header-message");
    var warningText = "Warning: Office.js is loaded outside of Office client";
    var winExt = window.external;
    var winExtGC = window.external.GetContext;
    var winHostInfo = window.external.GetHostInfo;
    var mb = Office.context.mailbox;
    var msg = "Window Ext: ";

    if (Office.context.mailbox) {
      alert("we are in outlook");
    } else {
      alert("we are NOT in outlook");
    }



    try
    {
        if(window.external && typeof window.external.GetContext !== "undefined")
             context = OSF.DDA._OsfControlContext = window.external.GetContext();
         else
         {
             msg = warningText;
         }
     }
     catch(e)
     {
         msg = warningText;
     }

    headMsg.textContent = msg;
    // End determine browser or outlook 




    // Office.context.mailbox.item.addFileAttachmentAsync(
    //   `https://webserver/picture.png`,
    //   "picture.png",
    //   { asyncContext: null },
    //   function(asyncResult) {
    //     if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    //       write(asyncResult.error.message);
    //     } else {
    //       // Get the ID of the attached file.
    //       var attachmentID = asyncResult.value;
    //       write("ID of added attachment: " + attachmentID);
    //     }
    //   }
    // );

    //  Office.context.mailbox.item.notificationMessages.replaceAsync("addin-message", {
    //     type: "informationalMessage",
    //     message: "Insert forms button was pressed",
    //     icon : "iconid",
    //     persistent: false
    //  });
  }

  // Writes to a div with id='message' on the page.
  function write(message) {
    var msg = document.getElementById("footer-message").textContent = message;
  }

})();
