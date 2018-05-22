/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

"use strict";

(function() {
  $(document).ready(function() {

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

    var inBtn = document.getElementById("insert-button");
    inBtn.onclick = insertSelectedForms;
    write("wired up insert button");

  });

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function(reason) {
    client = "could be office";
  };

  function run() {npm
    /**
     * Insert your Outlook code here
     */
  }

  function insertSelectedForms(obj) {
    //obj.currentTarget.style = "background: red";
    write("insert clicked ...");
  }

  // Writes to a div with id='message' on the page.
  function write(message) {
    var msg = document.getElementById("footer-message").textContent = message;
  }

})();
