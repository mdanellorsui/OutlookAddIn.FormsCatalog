// "use strict";

(function() {
  $(document).ready(function() {
    var inBtn = document.getElementById("insert-button");
    inBtn.onclick = insertSelectedForms;
    write("wired up insert button");
  });

  function insertSelectedForms(obj) {
    //obj.currentTarget.style = "background: red";
    write("insert clicked ...");
  }

  // Writes to a div with id='message' on the page.
  function write(message) {
    var msg = (document.getElementById("footer-message").textContent = message);
  }
})();
