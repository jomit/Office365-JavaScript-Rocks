
var O365JS = O365JS || {};

O365JS.core = (function () {
    "use strict";
    var clientId = '<client id here>...';
    var replyUrl = '<reply url here>...';

    function callService(options) {
        return $.Deferred(function (defer) {
            var type = "GET";
            if (options.type) {
                type = options.type;
            }
            $.ajax({
                url: options.apiUrl,
                type: type,
                dataType: "json",
                contentType: "application/json",
                beforeSend: function (xhr) {
                    xhr.setRequestHeader("Authorization", "Bearer " + options.accessToken);
                }
            }).done(function (data) {
                defer.resolve(data);
            }).fail(function (jqXHR, textStatus, errorThrown) {
                alert(jqXHR.responseText);
            });
        }).promise();
    }

    function authenticate(apiUrl) {
        var apiURI = URI(apiUrl);
        var url = 'https://login.windows.net/common/oauth2/authorize?' +
                   "response_type=" + encodeURI('token') + "&" +
                   "client_id=" + encodeURI(clientId) + "&" +
                   "resource=" + apiURI.protocol() + "://" + apiURI.authority() + "&" +
                   "redirect_uri=" + encodeURI(replyUrl);

        window.location = url;
    }   

    function getAccessTokenFromUrl() {
        var name = "access_token";
        var query = window.location.hash.substr(1);
        var vars = query.split("&");
        for (var i = 0; i < vars.length; i++) {
            var pair = vars[i].split("=");
            if (pair[0].toLowerCase() == name.toLowerCase()) { return pair[1]; }
        }
        return null;
    }

    return {
        callService: callService,
        authenticate: authenticate,
        getAccessTokenFromUrl: getAccessTokenFromUrl
    }
})();



