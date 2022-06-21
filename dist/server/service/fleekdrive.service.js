"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.FleekDriveService = void 0;
var tslib_1 = require("tslib");
var axios_1 = tslib_1.__importDefault(require("axios"));
var form_data_1 = tslib_1.__importDefault(require("form-data"));
var request_1 = tslib_1.__importDefault(require("request"));
var date_and_time_1 = tslib_1.__importDefault(require("date-and-time"));
var FleekDriveService = (function () {
    function FleekDriveService() {
        this.driveConfig = {
            fleekDriverId: process.env.FLEEK_DRIVER_ID,
            fleekDriverTk: process.env.FLEEK_DRIVER_TK,
            fleekDriverSpaceId: process.env.FLEEK_DRIVER_SPACE_ID,
            fleekDriverEndpoint: process.env.FLEEK_DRIVER_ENDPOINT,
            fleekDriverListId: process.env.FLEEK_DRIVER_LISTID
        };
    }
    FleekDriveService.prototype.authentication = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var _this = this;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4, axios_1.default
                            .get(this.driveConfig.fleekDriverEndpoint + "/common/Authentication.json", {
                            params: {
                                id: this.driveConfig.fleekDriverId,
                                tk: this.driveConfig.fleekDriverTk
                            },
                            withCredentials: true
                        })
                            .then(function (res) {
                            var e_1, _a;
                            var authCookies = res.headers["set-cookie"];
                            _this.authCookie = "";
                            if (authCookies) {
                                try {
                                    for (var authCookies_1 = tslib_1.__values(authCookies), authCookies_1_1 = authCookies_1.next(); !authCookies_1_1.done; authCookies_1_1 = authCookies_1.next()) {
                                        var cookie = authCookies_1_1.value;
                                        _this.authCookie += cookie + ";";
                                    }
                                }
                                catch (e_1_1) { e_1 = { error: e_1_1 }; }
                                finally {
                                    try {
                                        if (authCookies_1_1 && !authCookies_1_1.done && (_a = authCookies_1.return)) _a.call(authCookies_1);
                                    }
                                    finally { if (e_1) throw e_1.error; }
                                }
                            }
                        })
                            .catch(function (error) {
                            console.error(error);
                        })];
                    case 1:
                        _a.sent();
                        return [2];
                }
            });
        });
    };
    FleekDriveService.prototype.getRelatedListBySfdcId = function (recordId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var result;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        recordId = recordId ? recordId.substring(0, 15) : "0011m00000SvSEP";
                        return [4, axios_1.default
                                .get(this.driveConfig.fleekDriverEndpoint + "/contentsmanagement/relatedListRecordGetV2.json", {
                                params: {
                                    sf_objId: recordId,
                                    list_id: this.driveConfig.fleekDriverListId,
                                    thumb: 1,
                                    isDrilldown: false,
                                    isCurrent: 0,
                                    rows: 100,
                                    page: 1,
                                    sidx: "name",
                                    sord: "asc"
                                },
                                withCredentials: true,
                                headers: {
                                    Cookie: this.authCookie ? this.authCookie : ""
                                }
                            })
                                .then(function (res) {
                                result = res.data;
                            })
                                .catch(function (error) {
                                console.error(error);
                                result = null;
                            })];
                    case 1:
                        _a.sent();
                        return [2, new Promise(function (resolve) {
                                resolve(result);
                            })];
                }
            });
        });
    };
    FleekDriveService.prototype.getContentsRelationList = function (recordId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var result;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        recordId = recordId ? recordId : "0011m00000SvSEP";
                        return [4, axios_1.default
                                .get(this.driveConfig.fleekDriverEndpoint + "/contentsmanagement/contentsRelationListGet.json", {
                                params: {
                                    objid: recordId
                                },
                                withCredentials: true,
                                headers: {
                                    Cookie: this.authCookie ? this.authCookie : ""
                                }
                            })
                                .then(function (res) {
                                result = res.data;
                            })
                                .catch(function (error) {
                                console.error(error);
                                result = null;
                            })];
                    case 1:
                        _a.sent();
                        return [2, new Promise(function (resolve) {
                                resolve(result);
                            })];
                }
            });
        });
    };
    FleekDriveService.prototype.getSpaceContentsList = function (spaceId, searchStr) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var result;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4, axios_1.default
                            .get(this.driveConfig.fleekDriverEndpoint + "/contentsmanagement/spaceContentsListV2Get.json", {
                            params: {
                                asrt: "COMPANY",
                                spaceId: spaceId,
                                sord: "asc",
                                rows: 100,
                                sidx: "name",
                                page: 1,
                                searchstr: searchStr
                            },
                            withCredentials: true,
                            headers: {
                                Cookie: this.authCookie ? this.authCookie : ""
                            }
                        })
                            .then(function (res) {
                            result = res.data;
                        })
                            .catch(function (error) {
                            console.error(error);
                            result = null;
                        })];
                    case 1:
                        _a.sent();
                        return [2, new Promise(function (resolve) {
                                resolve(result);
                            })];
                }
            });
        });
    };
    FleekDriveService.prototype.getContentsList = function (spaceId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var result;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4, axios_1.default
                            .get(this.driveConfig.fleekDriverEndpoint + "/common/ContentsListGet.json", {
                            params: {
                                sord: "asc",
                                rows: 100,
                                sidx: "name",
                                page: 1,
                                ids: "[\"" + spaceId + "\"]",
                                includeSubSpace: 1
                            },
                            withCredentials: true,
                            headers: {
                                Cookie: this.authCookie ? this.authCookie : ""
                            }
                        })
                            .then(function (res) {
                            result = res.data;
                        })
                            .catch(function (error) {
                            console.error(error);
                            result = null;
                        })];
                    case 1:
                        _a.sent();
                        return [2, new Promise(function (resolve) {
                                resolve(result);
                            })];
                }
            });
        });
    };
    FleekDriveService.prototype.downloadContents = function (docId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var result, downLoadId;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4, axios_1.default
                            .get(this.driveConfig.fleekDriverEndpoint + "/contentsmanagement/SingleContentsDownloadApi.json", {
                            params: {
                                fileId: docId
                            },
                            withCredentials: true,
                            headers: {
                                Cookie: this.authCookie ? this.authCookie : ""
                            }
                        })
                            .then(function (res) {
                            downLoadId = res.data.id;
                        })
                            .catch(function (error) {
                            console.error(error);
                            result = null;
                        })];
                    case 1:
                        _a.sent();
                        if (!downLoadId)
                            return [2];
                        return [4, axios_1.default
                                .get(this.driveConfig.fleekDriverEndpoint + "/contentsmanagement/ContentsDownloadApi.json", {
                                params: {
                                    id: downLoadId
                                },
                                withCredentials: true,
                                headers: {
                                    Cookie: this.authCookie ? this.authCookie : ""
                                },
                                responseType: "stream"
                            })
                                .then(function (res) {
                                result = res.data;
                            })
                                .catch(function (error) {
                                console.error(error);
                                result = null;
                            })];
                    case 2:
                        _a.sent();
                        return [2, new Promise(function (resolve) {
                                resolve(result);
                            })];
                }
            });
        });
    };
    FleekDriveService.prototype.uploadFile = function (spaceId, fileName, fileBuffer) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var formData, today, result;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        formData = new form_data_1.default();
                        today = new Date();
                        formData.append("filename", fileName + "_" + today.getTime() + ".xlsx");
                        formData.append("spaceid", spaceId);
                        formData.append("checkNewVersion", "0");
                        formData.append("verUp", "");
                        formData.append("verUpType", "auto");
                        formData.append("versionkind", "0");
                        formData.append("file", fileBuffer);
                        return [4, axios_1.default
                                .post(this.driveConfig.fleekDriverEndpoint + "/contentsmanagement/ContentsAdd.json", formData, {
                                headers: tslib_1.__assign(tslib_1.__assign({}, formData.getHeaders()), { Cookie: this.authCookie ? this.authCookie : "" })
                            })
                                .then(function (res) {
                                result = res.data;
                            })
                                .catch(function (error) {
                                console.error(error);
                                result = null;
                            })];
                    case 1:
                        _a.sent();
                        return [2, new Promise(function (resolve) {
                                resolve(result);
                            })];
                }
            });
        });
    };
    FleekDriveService.prototype.uploadFile2 = function (spaceId, fileName, fileBuffer) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var httpOption, today, filename, promise;
            return tslib_1.__generator(this, function (_a) {
                httpOption = {
                    url: this.driveConfig.fleekDriverEndpoint + "/contentsmanagement/ContentsAdd.json",
                    headers: {
                        Cookie: this.authCookie ? this.authCookie : "",
                        "Cache-Control": "no-cache",
                        "Content-Type": "multipart/form-data"
                    }
                };
                today = new Date();
                filename = fileName + "_" + date_and_time_1.default.format(today, "YYYYMMDD") + ".xlsx";
                promise = new Promise(function (resolve) {
                    var r = request_1.default.post(httpOption, function (error, response, body) {
                        if (error) {
                            console.error(error);
                        }
                        resolve({
                            error: error,
                            response: response,
                            body: body,
                            fileName: filename
                        });
                    });
                    var formData = r.form();
                    formData.append("filename", filename);
                    formData.append("spaceid", spaceId);
                    formData.append("checkNewVersion", "0");
                    formData.append("verUp", "on");
                    formData.append("verUpType", "auto");
                    formData.append("versionkind", "0");
                    formData.append("file", fileBuffer, { filename: filename });
                    console.log(filename + "出力中");
                });
                return [2, promise];
            });
        });
    };
    return FleekDriveService;
}());
exports.FleekDriveService = FleekDriveService;
//# sourceMappingURL=fleekdrive.service.js.map