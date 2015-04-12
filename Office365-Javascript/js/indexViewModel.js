/// <reference path="core.js" />
/// <reference path="lib/knockout-3.3.0.js" />
/// <reference path="lib/jquery-2.1.3.min.js" />

(function (O365JS) {
    "use strict";
    O365JS.indexViewModel = (function () {
        var accessToken = ko.observable();
        var myDocuments = ko.observableArray();
        var apiUrl = 'https://<your o365 domain here>-my.sharepoint.com/_api/v1.0/me/files';

        function initialize() {
            var token = O365JS.core.getAccessTokenFromUrl();
            if (token != null) {
                accessToken(token);
            }
        }

        function authenticate() {
            var token = O365JS.core.getAccessTokenFromUrl();
            if (token == null) {
                O365JS.core.authenticate(apiUrl);
            }
            else {
                accessToken(token);
            }
        }

        function getDocuments() {
            O365JS.core.callService({ apiUrl: apiUrl , accessToken: accessToken()})
                .then(function (data) {
                    var files = ko.utils.arrayFilter(data.value, function (docItem) {
                        return docItem.type == "File";
                    });
                    myDocuments(files);
                });
        }

        return {
            initialize: initialize,
            authenticate: authenticate,
            accessToken: accessToken,
            myDocuments: myDocuments,
            getDocuments: getDocuments
        }
    })();
    var indexView = document.getElementById("indexViewModel");
    if (indexView !== null) {
        O365JS.indexViewModel.initialize();
        ko.applyBindings(O365JS.indexViewModel, indexView);
    }
})(O365JS);