// Configuration: Obtain Slack web API token at https://api.slack.com/web
// Caution: You can't count files in DM or Private Files
var API_TOKEN = PropertiesService.getScriptProperties().getProperty('api_token'),
    STORAGE_LIMIT = PropertiesService.getScriptProperties().getProperty('storage_limit'),
    FROM_DAYS = PropertiesService.getScriptProperties().getProperty('from_days'),
    TO_DAYS = PropertiesService.getScriptProperties().getProperty('to_days'),
    FILES_LIST_OPTION_USER = PropertiesService.getScriptProperties().getProperty('user'),
    FILES_LIST_OPTION_CHANNEL = PropertiesService.getScriptProperties().getProperty('channel'),
    FILES_LIST_OPTION_TS_FROM = PropertiesService.getScriptProperties().getProperty('ts_from'),
    FILES_LIST_OPTION_TS_TO = PropertiesService.getScriptProperties().getProperty('ts_to'),
    FILES_LIST_OPTION_TYPES = PropertiesService.getScriptProperties().getProperty('types'),
    FILES_LIST_OPTION_COUNT = PropertiesService.getScriptProperties().getProperty('count'),
    FILES_LIST_OPTION_PAGE = PropertiesService.getScriptProperties().getProperty('page');

if (!API_TOKEN) {
    throw 'You should set "slack_api_token" property from [File] > [Project properties] > [Script properties]';
}
if (!STORAGE_LIMIT) {
    throw 'You should set "STORAGE_LIMIT" property from [File] > [Project properties] > [Script properties]';
}
if (!FILES_LIST_OPTION_USER) {
    FILES_LIST_OPTION_USER = '';
}
if (!FILES_LIST_OPTION_CHANNEL) {
    FILES_LIST_OPTION_CHANNEL = '';
}
if (!FILES_LIST_OPTION_TS_FROM) {
    FILES_LIST_OPTION_TS_FROM = '';
}
if (!FILES_LIST_OPTION_TS_TO) {
    FILES_LIST_OPTION_TS_TO = '';
}
if (!FILES_LIST_OPTION_TYPES) {
    FILES_LIST_OPTION_TYPES = '';
}
if (!FILES_LIST_OPTION_COUNT) {
    FILES_LIST_OPTION_COUNT = '';
}
if (!FILES_LIST_OPTION_PAGE) {
    FILES_LIST_OPTION_PAGE = '';
}

function DeleteFiles() {
    var slackFilesDeleter = new SlackFilesDeleter();
    slackFilesDeleter.run();
}

var SlackFilesDeleter = (function () {

    var SlackFilesDeleter = function () {
        this.allFileSize = 0;
        this.deleteSize = 0;
        this.channelArray = {};
        if (!FILES_LIST_OPTION_TS_FROM) {
            if (FROM_DAYS) {
                FILES_LIST_OPTION_TS_FROM = this.getOfBeforeAfterDays(new Date(), FROM_DAYS);
            }
        }
        if (!FILES_LIST_OPTION_TS_TO) {
            if (TO_DAYS) {
                FILES_LIST_OPTION_TS_TO = this.getOfBeforeAfterDays(new Date(), TO_DAYS);
            }
        }
    };

    SlackFilesDeleter.prototype.run = function () {
        Logger.log("getTotal");
        var total = this.getTotal(),
            pages = this.calcPages(total),
            responseDataArray,
            filesListURLAddParam;

        Logger.log("isOverStorage");
        if (this.isOverStorage(pages, STORAGE_LIMIT)) {
            Logger.log("Your storage is full");

            this.deleteSize = this.allFileSize - STORAGE_LIMIT * Math.pow(10, 9);
            Logger.log("All File : " + this.allFileSize / Math.pow(10, 6) + " MB");
            Logger.log("Unnecessary Files : " + this.deleteSize / Math.pow(10, 6) + " MB");

            Logger.log("get all file to check pages");
            responseDataArray = this.requestSlackAPI(this.getFilesListURL());
            for (var i = responseDataArray['paging'].pages; i > 0 && this.deleteSize > 0; i--) {
                FILES_LIST_OPTION_PAGE = i;
                filesListURLAddParam = this.getFilesListURL();
                Logger.log("get delete files");
                responseDataArray = this.requestSlackAPI(filesListURLAddParam);
                this.deleteFiles(this.checkPinnedForFiles(responseDataArray));
            }
        } else {
            Logger.log("All File : " + this.allFileSize / Math.pow(10, 6) + " MB");
            Logger.log("Your storage is not full");
        }
    };

    SlackFilesDeleter.prototype.requestSlackAPI = function (URL) {
        var response,
            responseDataArray;

        response = UrlFetchApp.fetch(URL);
        responseDataArray = JSON.parse(response);
        Logger.log("==> GET " + URL);
        if (responseDataArray.error) {
            throw "GET : " + responseDataArray.error;
        }

        return responseDataArray;
    };

    SlackFilesDeleter.prototype.getTotal = function () {
        var filesListURL = 'https://slack.com/api/files.list',
            filesListData = {
                'token': API_TOKEN,
                'pretty': 1 // 値0だとjsonの形にparseされず、それ以外ならparseされる模様
            },
            filesListURLAddParam = filesListURL + this.encodeHTMLForm(filesListData),
            responseDataArray = this.requestSlackAPI(filesListURLAddParam);

        return responseDataArray['paging']['total'];
    };

    SlackFilesDeleter.prototype.calcPages = function (total) {
        var pages;

        if (FILES_LIST_OPTION_COUNT === '') {
            if (total > 1000) {
                FILES_LIST_OPTION_COUNT = 1000;
                pages = Math.ceil(total / 1000);

                return pages;
            } else {
                FILES_LIST_OPTION_COUNT = total;
                return 1;
            }
        } else {
            return FILES_LIST_OPTION_COUNT;
        }
    };

    SlackFilesDeleter.prototype.checkChannelsID = function () {
        var channelsListURL = 'https://slack.com/api/channels.list',
            channelsListData = {
                'token': API_TOKEN,
                'pretty': 1
            },
            channelsListAddParam,
            response,
            channelsIDList = [];

        channelsListAddParam = channelsListURL + this.encodeHTMLForm(channelsListData);
        response = this.requestSlackAPI(channelsListAddParam);
        for (var i = 0; i < response['channels'].length; i++) {
            channelsIDList.push([response['channels'][i].id, response['channels'][i].name, 0]);
        }

        return channelsIDList;
    };

    SlackFilesDeleter.prototype.countChannelSizes = function (responseDataArray) {
        //DMのサイズはカウントされないため全体のファイルサイズより少ない可能性あり
        for (var i = 0; i < this.channelArray.length; i++) {
            for (var j = 0; j < responseDataArray['files'].length; j++) {
                if (this.channelArray[i][0] == responseDataArray['files'][j].channels) {
                    this.channelArray[i][2] += responseDataArray['files'][j].size;
                }
            }
        }
    };

    SlackFilesDeleter.prototype.createLogs = function (array) {
        var folder = this.getLogsFolder(),
            spreadsheet = SpreadsheetApp.create(Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd (E) HH:mm:ss")),
            sum = 0;

        folder.addFile(DriveApp.getFileById(spreadsheet.getId()));
        DriveApp.getRootFolder().removeFile(DriveApp.getFileById(spreadsheet.getId()));

        array.sort(function (a, b) {
            if (a[2] > b[2]) return -1;
            if (a[2] < b[2]) return 1;
            return 0;
        });

        spreadsheet.appendRow(["ID", "チャンネル", "サイズ(MB)"]);
        spreadsheet.setFrozenRows(1);
        for (var i = 0; i < array.length; i++) {
            spreadsheet.appendRow([array[i][0], array[i][1], array[i][2]]);
            sum += array[i][2];
        }
        spreadsheet.appendRow(['', "その他(DM等)", this.allFileSize / Math.pow(10, 6) - sum]);
        spreadsheet.appendRow(['', "チャンネル合計", sum]);
        spreadsheet.appendRow(['', "全体合計", this.allFileSize / Math.pow(10, 6)]);
        spreadsheet.insertRowBefore(spreadsheet.getLastRow() - 1);
    };

    SlackFilesDeleter.prototype.getLogsFolder = function () {
        var folder = DriveApp.getRootFolder(),
            name = "Slack File Sizes Logs",
            it = folder.getFoldersByName(name);

        if (it.hasNext()) {
            folder = it.next();
        } else {
            folder = folder.createFolder(name);
        }
        return folder;
    };

    SlackFilesDeleter.prototype.isOverStorage = function (pages, limitArg) {
        var filesListURL,
            filesListData,
            filesListURLAddParam,
            limit = limitArg * Math.pow(10, 9);

        this.allFileSize = 0;

        filesListURL = 'https://slack.com/api/files.list';
        filesListData = {
            'token': API_TOKEN,
            'count': FILES_LIST_OPTION_COUNT,
            'page': 1,
            'pretty': 1 // 値0だとjsonの形にparseされず、それ以外ならparseされる模様
        };
        filesListURLAddParam = filesListURL + this.encodeHTMLForm(filesListData);

        this.channelArray = this.checkChannelsID();

        for (var i = 1; i <= pages; i++) {
            filesListData['page'] = i;
            filesListURLAddParam = filesListURL + this.encodeHTMLForm(filesListData);
            this.addSizes(this.requestSlackAPI(filesListURLAddParam));
            this.countChannelSizes(this.requestSlackAPI(filesListURLAddParam));
        }

        for (var i = 0; i < this.channelArray.length; i++) {
            this.channelArray[i][2] /= Math.pow(10, 6);
        }

        this.createLogs(this.channelArray);

        if (this.allFileSize >= limit) {
            return true;
        } else {
            return false;
        }
    };

    SlackFilesDeleter.prototype.addSizes = function (responseDataArray) {
        var sizeCount = 0;

        for (var i = 0; i < responseDataArray['files'].length; i++) {
            sizeCount += responseDataArray['files'][i].size;
        }

        this.allFileSize += sizeCount;
    };

    SlackFilesDeleter.prototype.deleteFiles = function (checkedfileArray) {
        var filesDeleteURL = 'https://slack.com/api/files.delete',
            filesDeleteData = {
                'token': API_TOKEN,
                'file': 0,
                'pretty': 1
            },
            filesDeleteURLAddParam;

        for (var i = checkedfileArray.length - 1; i >= 0 && this.deleteSize > 0; i--) {
            filesDeleteData['file'] = checkedfileArray[i].id;
            filesDeleteURLAddParam = filesDeleteURL + this.encodeHTMLForm(filesDeleteData);
            this.requestSlackAPI(filesDeleteURLAddParam);
            this.deleteSize -= checkedfileArray[i].size;
            Logger.log("size : " + checkedfileArray[i].size + "   id : " + checkedfileArray[i].id + "   left : " + this.deleteSize / Math.pow(10, 6) + " MB");
        }
    };

    SlackFilesDeleter.prototype.checkPinnedForFiles = function (responseDataArray) {
        var responseFiles;
        responseFiles = JSON.stringify(responseDataArray['files']); // JSON文字列化
        responseFiles = JSON.parse(responseFiles);

        for (var i = responseFiles.length - 1; i >= 0; i--) {
            if (responseFiles[i].pinned_to) {
                responseFiles.splice(i, 1);
            }
        }

        return responseFiles;
    };

    SlackFilesDeleter.prototype.getFilesListURL = function () {
        var filesListURL = 'https://slack.com/api/files.list',
            filesListData = {
                'token': API_TOKEN,
                'user': FILES_LIST_OPTION_USER,
                'channel': FILES_LIST_OPTION_CHANNEL,
                'ts_from': FILES_LIST_OPTION_TS_FROM,
                'ts_to': FILES_LIST_OPTION_TS_TO,
                'types': FILES_LIST_OPTION_TYPES,
                'count': FILES_LIST_OPTION_COUNT,
                'page': FILES_LIST_OPTION_PAGE,
                'pretty': 1 // 値0だとjsonの形にparseされず、それ以外ならparseされる模様
            },
            filesListURLAddParam = filesListURL + this.encodeHTMLForm(filesListData);

        return filesListURLAddParam;
    };

    SlackFilesDeleter.prototype.getOfBeforeAfterDays = function (dateObj, number) {
        var result;

        result = dateObj.getTime() + number * 24 * 60 * 60 * 1000 * -1;
        result = Math.floor(result / 1000);

        return result;
    };

    SlackFilesDeleter.prototype.encodeHTMLForm = function (data) {
        var params = [],
            paramsStr = '';

        for (var name in data) {
            var value = data[name];
            if (value !== '') {
                var param = encodeURIComponent(name) + '=' + encodeURIComponent(value);
                params.push(param);
            }
        }

        paramsStr = params.join('&').replace(/%20/g, '+');
        paramsStr = '?' + paramsStr;

        return paramsStr;
    };

    return SlackFilesDeleter;
}());
