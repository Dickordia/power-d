(function () {
"use strict";
    function updateControls(index) {
      const aData = JSON.parse(localStorage.getItem("data"))[index]
      document.getElementById("index-box").value = index;
      document.getElementById("value-box").value = aData.value;
      document.getElementById("color-box").value = aData.color;
    }

    Office.onReady()
        .then(function() {
            updateControls(0)
            document.getElementById("index-box").onchange = updateIndexToParentPage;
            document.getElementById("color-box").onchange = sendColorToParentPage;
            document.getElementById("update-button").onclick = sendUpdateToParentPage;
            document.getElementById("close-button").onclick = sendStringToParentPage;
        });

        function sendStringToParentPage() {
          Office.context.ui.messageParent("");
        }

        function sendColorToParentPage() {
          var aIndex = document.getElementById("index-box").value;
          var aColor = document.getElementById("color-box").value;
          var aRes = JSON.stringify({index:aIndex, color: aColor})
          Office.context.ui.messageParent(aRes);
        }

        function updateIndexToParentPage() {
          var aIndex = document.getElementById("index-box").value;
          updateControls(aIndex)
        }

        function sendUpdateToParentPage() {
          var aIndex = document.getElementById("index-box").value;
          var aValue = document.getElementById("value-box").value;
          var aColor = document.getElementById("color-box").value;
          var aRes = JSON.stringify({index:aIndex, value: aValue, color: aColor})
          Office.context.ui.messageParent(aRes);
        }
}());
