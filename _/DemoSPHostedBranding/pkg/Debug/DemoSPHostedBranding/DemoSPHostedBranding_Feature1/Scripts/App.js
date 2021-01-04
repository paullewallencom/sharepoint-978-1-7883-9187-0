'use strict';

ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

function initializePage()
{
    var context = SP.ClientContext.get_current();
    var user = context.get_web().get_currentUser();

    // This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
    $(document).ready(function () {
        $('#applyCustomization').click(function () {
            outputMessage("Starting...");

            var hostWebUrl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
            var appContectSite = new SP.AppContextSite(context, hostWebUrl);
            var hostWeb = appContectSite.get_web();
           
            BrandingHelper.UploadFile(context, hostWeb, hostWebUrl, "Branding/Files/logo.png", "DemoBrandingApp", 116);
            BrandingHelper.UploadFile(context, hostWeb, hostWebUrl,
                "Branding/MasterPages/default.master", "DemoBrandingApp", 116).done(function (fileUrl) {
                    BrandingHelper.ActivateMasterPage(context, hostWeb, fileUrl).done(function (res) {
                        if (res) {
                            BrandingHelper.UploadPageLayout(context, hostWeb, hostWebUrl, "Branding/PageLayouts/ArticleLeftCustom.aspx", "DemoBrandingApp").done(function (uploadResult) {
                                if (uploadResult) {
                                    outputMessage("Done, enabled.");
                                } else {
                                    outputMessage("Error occurred.");
                                }
                            });
                            outputMessage("Done, enabled.");
                        } else {
                            outputMessage("Error occurred.");
                        }
                    });
            });
            return false;
        });

        $('#disableCustomization').click(function () {
            outputMessage("Starting...");

            var hostWebUrl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
            var appContectSite = new SP.AppContextSite(context, hostWebUrl);
            var hostWeb = appContectSite.get_web();
            
            BrandingHelper.DeactivateMasterPage(context, hostWeb).done(function (msg) {
                outputMessage(msg);
            })

            return false;
        });
    });

    function getQueryStringParameter(param) {
        var params = document.URL.split("?")[1].split("&");
        var strParams = "";
        for (var i = 0; i < params.length; i = i + 1) {
            var singleParam = params[i].split("=");
            if (singleParam[0] == param) {
                return singleParam[1];
            }
        }
    }

    function outputMessage(msg) { $('#message').text(msg); }
}
