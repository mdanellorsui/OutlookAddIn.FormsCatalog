/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';
var client = 'unknown';

(function () {

  $(document).ready(function () {
    //  $('#run').click(run);

      // Office.context.mailbox.item.notificationMessages.replaceAsync("progress1", {
      //   type: "errorMessage",
      //   message : "Foo Test"
      //   });

      var inBtn = document.getElementById("insert-button");
      inBtn.onclick = insertSelectedForms;
      //      $('#insert-button').on('click', insertSelectedForms);
  });

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    client = 'could be office';
  };

  function run() {
    /**
     * Insert your Outlook code here
     */
  }

  function insertSelectedForms(obj) {

    obj.currentTarget.innerText = "foo";
    obj.currentTarget.style = "background: red";

    // Office.context.mailbox.item.notificationMessages.replaceAsync("addin-message", {
    //   type: "informationalMessage",
    //   message: "Insert forms button was pressed",
    //   icon : "iconid",
    //   persistent: false
    // });
  }


})();