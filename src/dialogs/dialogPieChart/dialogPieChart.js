(function () {
"use strict";

    Office.onReady()
        .then(function() {
            document.getElementById("index-box").value = 0;
            document.getElementById("value-box").value = "0.0";
            document.getElementById("update-button").onclick = sendUpdateToParentPage;
            document.getElementById("close-button").onclick = sendStringToParentPage;
        });

        function sendStringToParentPage() {
          Office.context.ui.messageParent("");
        }

        function sendUpdateToParentPage() {
          var aIndex = document.getElementById("index-box").value;
          var aValue = document.getElementById("value-box").value;
          var aRes = JSON.stringify({index:aIndex, value: aValue})
          Office.context.ui.messageParent(aRes);
        }
}());
