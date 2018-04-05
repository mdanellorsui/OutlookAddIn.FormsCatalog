/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

(function () {

  $(document).ready(function () {
    //  $('#run').click(run);

      // Office.context.mailbox.item.notificationMessages.replaceAsync("progress1", {
      //   type: "errorMessage",
      //   message : "Foo Test"
      //   });
      var inBtn = document.getElementById("insert-button");
      inBtn.innerText = "foo";
      //$("#insert-button").style="background:red";  

      //$('#insert-button').on('click', insertSelectedForms);


  });

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {

  };

  function run() {
    /**
     * Insert your Outlook code here
     */
  }

  function insertSelectedForms() {
    Office.context.mailbox.item.notificationMessages.replaceAsync("addin-message", {
      type: "informationalMessage",
      message: "Insert forms button was pressed",
      icon : "iconid",
      persistent: false
    });
  }


})();