'use strict';

SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
    
    String.prototype.trimStart = function (c) {
        if (this.length == 0)
            return this;
        c = c ? c : ' ';
        var i = 0;
        var val = 0;
        for (; this.charAt(i) == c && i < this.length; i++);
        return this.substring(i);
    }
    String.prototype.trimEnd = function (c) {
        c = c ? c : ' ';
        var i = this.length - 1;
        for (; i >= 0 && this.charAt(i) == c; i--);
        return this.substring(0, i + 1);
    }
    String.prototype.trim = function (c) {
        return this.trimStart(c).trimEnd(c);
    }

    SP.ClientContext.prototype.executeQuery = function () {
        var deferred = $.Deferred();
        this.executeQueryAsync(
            function () { deferred.resolve(arguments); },
            function () { deferred.reject(arguments); }
        );
        return deferred.promise();
    };

    window.BrandingHelper = window.BrandingHelper || {};
    BrandingHelper.trimChars = ['/'];

    BrandingHelper.UploadFile = function (context, hostWeb, hostWebUrl, fileName, folderName, listInternalName) {
        var deferred = $.Deferred();
        var appWeb = context.get_web();
        var gallery = null;
        var rootFolder = null;
        if (listInternalName == 116) {
            gallery = hostWeb.getCatalog(116);
            rootFolder = gallery.get_rootFolder();
            context.load(gallery);
            context.load(rootFolder);
        }
        context.load(appWeb);
        context.executeQuery().done(function () {
            var serverRelativeAppUrl = appWeb.get_serverRelativeUrl().trimEnd(BrandingHelper.trimChars);
            var serverRelativeHostWebUrl = serverRelativeAppUrl.substring(0, serverRelativeAppUrl.lastIndexOf('/'));
            var startingUrl = null;
            if (listInternalName == 116) {
                startingUrl = rootFolder.get_serverRelativeUrl().trimEnd(BrandingHelper.trimChars) + "/";
            } else {
                startingUrl = serverRelativeHostWebUrl + "/" + listInternalName + "/";
            }
            console.log(startingUrl);
            // folder/sub1/sub2
            BrandingHelper.EnsureFolders(context, hostWeb,  startingUrl, folderName, fileName).done(function (folder) {
                BrandingHelper.AddFile(context, serverRelativeAppUrl,
                    hostWeb, hostWebUrl, fileName, folder.get_serverRelativeUrl()).done(function (fileUrl) {
                        if (fileUrl) {
                            deferred.resolve(fileUrl);
                        } else {
                            deferred.resolve(null);
                        }
                    });
                });
        }).fail(function () {
            deferred.resolve(null);
        });

        return deferred.promise();
    };

    BrandingHelper.EnsureFolder = function (clientContext, hostWeb, listUrl, folderUrl, parentFolder) {
        var deferredFolderResult = new $.Deferred();
        var folder = null;
        var folderServerRelativeUrl = parentFolder == null ? listUrl.trimEnd(BrandingHelper.trimChars) + "/" + folderUrl : parentFolder.get_serverRelativeUrl().trimEnd(BrandingHelper.trimChars) + "/" + folderUrl;

        folder = hostWeb.getFolderByServerRelativeUrl(folderServerRelativeUrl);
        clientContext.load(folder);

        clientContext.executeQuery().done(function () {
            deferredFolderResult.resolve(folder);
        }).fail(function (args) {
            var lists = hostWeb.get_lists();
            clientContext.load(lists, 'Include(DefaultViewUrl)');
            clientContext.executeQuery().done(function () {
                var list = null;
                for (var i = 0; i < lists.get_count() ; i++) {
                    if (lists.getItemAtIndex(i).get_defaultViewUrl().indexOf(listUrl) != -1) {
                        list = lists.getItemAtIndex(i);
                        break;
                    }
                }
                if (list != null) {
                    clientContext.load(list);
                    if (parentFolder == null) {
                        parentFolder = list.get_rootFolder();
                    }
                    folder = parentFolder.get_folders().add(folderUrl);
                    clientContext.load(folder);
                    clientContext.executeQuery().done(function () {
                        deferredFolderResult.resolve(folder);
                    }).fail(function (args) {
                        console.log(args[1].get_message());
                        deferredFolderResult.resolve(null);
                    });
                } else {
                    deferredFolderResult.resolve(null);
                }
            }).fail(function (args) {
                console.log(args[1].get_message());
                deferredFolderResult.resolve(null);
            });
        });

        return deferredFolderResult.promise();
    }

    BrandingHelper.EnsureFolders = function (clientContext, hostWeb, filePath, fileFolder, fileName) {
        var deferredResult = $.Deferred();
        BrandingHelper.EnsureFolder(clientContext, hostWeb, filePath, fileFolder, null).done(function (f) {
            var folder = f;
            console.log(folder);

            if (fileName == null || fileName == '') {
                deferredResult.resolve(folder);
            }

            // path to file may contain folders, we need to ensure they exist, too
            var fileNameSplitted = fileName.split('/');
            var fileNameSplittedFoldersOnly = fileNameSplitted.splice(0, fileNameSplitted.length - 1);
            var itemsProcessed = 0;
            async.eachSeries(fileNameSplittedFoldersOnly, function (folderName, callback) {
                folderPromise = BrandingHelper.EnsureFolder(clientContext, hostWeb, filePath, folderName, folder);
                folderPromise.then(function (f) {
                    folder = f;
                    itemsProcessed++;
                    callback();
                });
            }, function (err) {
                if (err) {
                    console.log(err);
                    deferredResult.resolve(null);
                }

                if (itemsProcessed == fileNameSplittedFoldersOnly.length) {
                    deferredResult.resolve(folder);
                }
            });
        }).fail(function(args) {
            console.log(args[1].get_message());
            deferredResult.resolve(null);
        });

        return deferredResult.promise();
    }

    BrandingHelper.GetAppWebFile = function (context, appWebServerRelativeUrl, fileUrl) {
        var deferred = $.Deferred();

        var urlToFile = window.location.protocol + "//" + window.location.hostname + "/" + appWebServerRelativeUrl + "/" + fileUrl;
        console.log(urlToFile);

        var oReq = new XMLHttpRequest();
        oReq.open("GET", urlToFile, true);
        oReq.responseType = "arraybuffer";

        oReq.onload = function (oEvent) {
            var arrayBuffer = oReq.response;
            if (arrayBuffer) {
                var byteArray = new Uint8Array(arrayBuffer);
                deferred.resolve(byteArray);
            } else {
                deferred.resolve(null);
            }
        };

        oReq.send(null);
        return deferred.promise();
    }

    BrandingHelper.AddFile = function (context, appWebServerRelativeUrl, hostWeb, hostWebUrl, fileName, hostWebFolderUrl) {
        var deferred = $.Deferred();
        BrandingHelper.GetAppWebFile(context, appWebServerRelativeUrl, fileName).done(function (appFile) {
            if (appFile == null) {
                var errMsg = "Could not find file " + fileName + " in app web";
                console.log(errMsg);
                deferred.resolve(false);
                return;
            }

            var pureFileName = fileName;
            if (fileName.indexOf('/') != -1) {
                pureFileName = fileName.substr(fileName.lastIndexOf('/') + 1);
            }
            var appWebUrl = window.location.protocol + "//" + window.location.hostname + "/" + appWebServerRelativeUrl;
            var fileCollectionEndpoint = String.format(
                "{0}/_api/sp.appcontextsite(@target)/web/getfolderbyserverrelativeurl('{1}')/files" +
                "/add(overwrite=true, url='{2}')?@target='{3}'",
                 appWebUrl, hostWebFolderUrl, pureFileName, hostWebUrl);

            // Send the request and return the response.
            // This call returns the SharePoint file.
            var uploadPromise = $.ajax({
                url: fileCollectionEndpoint,
                type: "POST",
                data: appFile,
                processData: false,
                headers: {
                    "accept": "application/json;odata=verbose",
                    "X-RequestDigest": jQuery("#__REQUESTDIGEST").val(),
                    "content-length": appFile.byteLength
                }
            });

            uploadPromise.done(function () {
                console.log("Added file to this url: " + hostWebUrl + hostWebFolderUrl + "/" + pureFileName);
                deferred.resolve(hostWebUrl + hostWebFolderUrl + "/" + pureFileName); 
            }).fail(function () {
                deferred.resolve(false);
            })
        });
        return deferred.promise();
    }

    BrandingHelper.ActivateMasterPage = function (context, hostWeb, masterPageFileUrl) {
        var deferred = $.Deferred();
        var tmpTag = document.createElement('a');
        tmpTag.href = masterPageFileUrl;
        var properUrl = masterPageFileUrl.replace(tmpTag.protocol + "//", "")
            .replace(tmpTag.hostname, "");

        hostWeb.set_masterUrl(properUrl);
        hostWeb.set_customMasterUrl(properUrl);
        hostWeb.update();
        context.executeQuery().done(function () {
            deferred.resolve(true);
        }).fail(function (args) {
            console.log(args[1].get_message());
            deferred.resolve(false);
        });
        return deferred.promise();
    }

    BrandingHelper.DeactivateMasterPage = function (context, hostWeb) {
        var deferred = $.Deferred();
        var properUrl = "/_catalogs/masterpage/seattle.master";
        hostWeb.set_masterUrl(properUrl);
        hostWeb.set_customMasterUrl(properUrl);
        hostWeb.update();
        context.executeQuery().done(function () {
            deferred.resolve("Done, deactivated.");
        }).fail(function (args) {
            deferred.resolve(args[1].get_message());
        });
        return deferred.promise();
    }

    BrandingHelper.UploadPageLayout = function (context, hostWeb, hostWebUrl, fileName, folderName) {
        var deferred = $.Deferred();
        var lists = hostWeb.get_lists();
        var gallery = hostWeb.getCatalog(116);
        var rootFolder = gallery.get_rootFolder();
        var appWeb = context.get_web();
        context.load(appWeb);
        context.load(lists);
        context.load(gallery);
        context.load(rootFolder);

        context.executeQuery().done(function () {
            var masterPath = rootFolder.get_serverRelativeUrl().trimEnd(BrandingHelper.trimChars) + "/";
            BrandingHelper.EnsureFolder(context, hostWeb, masterPath, folderName, null).done(function (folderResult) {
                if (folderResult) {
                    var serverRelativeAppUrl = appWeb.get_serverRelativeUrl().trimEnd(BrandingHelper.trimChars);
                    var serverRelativeHostWebUrl = serverRelativeAppUrl.substring(0, serverRelativeAppUrl.lastIndexOf('/'));
                    BrandingHelper.AddFile(context, serverRelativeAppUrl,
                       hostWeb, hostWebUrl, fileName, folderResult.get_serverRelativeUrl()).done(function (fileUrl) {
                           if (fileUrl) {
                               deferred.resolve(true);
                           } else {
                               deferred.resolve(false);
                           }
                       });
                    deferred.resolve(true);
                } else {
                    deferred.resolve(false);
                }
            });
        }).fail(function (args) {
            console.log(args[1].get_message());
            deferred.resolve(false);
        });
        return deferred.promise();
    }
});