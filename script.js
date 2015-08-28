/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See License.txt in the project root.
*/

// The initialize function must be run each time a new page is loaded
Office.initialize = function (reason) {
	$(document).ready(function () {
		$('#set-data').click(writeText);
	});

	 //UI Components init
     $(".ms-Pivot").Pivot();
     $(".ms-SearchBox").SearchBox();
     $(".ms-Dropdown").Dropdown();
};

// Reads data from current document selection and displays a notification
function writeText() {
    Office.context.document.setSelectedDataAsync("Citation goes here",
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed"){
            	$('#display-data').text("Failure" + error.message);
            }
            else
            {
            	$('#display-data').text("Done");
            }
        });
}

