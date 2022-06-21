"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.FilesController = void 0;
var tslib_1 = require("tslib");
var express_1 = require("express");
var multer_1 = tslib_1.__importDefault(require("multer"));
var fleekdrive_service_1 = require("../service/fleekdrive.service");
var excel_service_1 = require("../service/excel.service");
var sfdc_service_1 = require("../service/sfdc.service");
var weeklyReport_service_1 = require("../service/weeklyReport.service");
var analysisReport_service_1 = require("../service/analysisReport.service");
var FilesController = (function () {
    function FilesController() {
        var _this = this;
        this.upload = multer_1.default();
        this.excelService = new excel_service_1.ExcelService();
        this.flkDriveService = new fleekdrive_service_1.FleekDriveService();
        this.getExcel = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var fileBuffer, fileName;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.excelService.addSheet("work");
                        return [4, this.excelService.getFileBuffer()];
                    case 1:
                        fileBuffer = _a.sent();
                        fileName = "demo.xlsx";
                        res.set("Content-disposition", "attachment; filename=" + fileName);
                        res.set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                        res.end(fileBuffer);
                        return [2];
                }
            });
        }); };
        this.getExcel2 = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                res.send("ok");
                return [2];
            });
        }); };
        this.getStatustable = function (req, res, next) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var httpRes, pjcode, statustableConfig, outputPatternMap, contractPatternMap, acc, statustableSpaceId, filesRes, tempFileName, statusTableTemplateId, _a, _b, file, error_1;
            var e_1, _c;
            return tslib_1.__generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        httpRes = { status: 0, data: null, msg: "ファイル作成中" };
                        _d.label = 1;
                    case 1:
                        _d.trys.push([1, 5, , 6]);
                        pjcode = req.params.pjcode;
                        statustableConfig = {
                            projectId: req.body.projectId,
                            projectName: "",
                            buildId: req.body.buildId,
                            outputPattern: req.body.outputPattern,
                            contractPattern: req.body.contractPattern,
                            customFiledSet: req.body.customFiledSet,
                            fileSuffix: "",
                            userId: req.body.userId
                        };
                        console.info(statustableConfig);
                        outputPatternMap = new Map();
                        outputPatternMap.set("contract", "契約パターン");
                        outputPatternMap.set("lottery", "抽選パターン");
                        outputPatternMap.set("request", "要望パターン");
                        contractPatternMap = new Map();
                        contractPatternMap.set("butterfly", "部屋属性+蝶々");
                        contractPatternMap.set("info", "部屋属性+個人属性");
                        contractPatternMap.set("btInfo", "部屋属性+個人属性+蝶々");
                        statustableConfig.fileSuffix = outputPatternMap.get(statustableConfig.outputPattern);
                        if (statustableConfig.contractPattern && statustableConfig.outputPattern === "contract") {
                            statustableConfig.fileSuffix += "-" + contractPatternMap.get(statustableConfig.contractPattern);
                        }
                        statustableConfig.fileSuffix = "(" + statustableConfig.fileSuffix + ")";
                        return [4, this.getAccount(statustableConfig.projectId, pjcode)];
                    case 2:
                        acc = _d.sent();
                        if (!acc) {
                            httpRes = { status: 1, data: null, msg: "プロジェクトが存在しません" };
                            return [2, res.json(httpRes)];
                        }
                        statustableConfig.projectName = acc.Name;
                        return [4, this.getSubSpaceId(acc.Id, "状況表")];
                    case 3:
                        statustableSpaceId = _d.sent();
                        if (!statustableSpaceId) {
                            httpRes = { status: 1, data: null, msg: "FleekDriveに「状況表」フォルダがありません。" };
                            return [2, res.json(httpRes)];
                        }
                        return [4, this.flkDriveService.getContentsList(statustableSpaceId)];
                    case 4:
                        filesRes = _d.sent();
                        tempFileName = "\u72B6\u6CC1\u8868\u30C6\u30F3\u30D7\u30EC\u30FC\u30C8" + statustableConfig.fileSuffix + ".xlsx";
                        statusTableTemplateId = void 0;
                        try {
                            for (_a = tslib_1.__values(filesRes.rows), _b = _a.next(); !_b.done; _b = _a.next()) {
                                file = _b.value;
                                if (file.name === tempFileName) {
                                    console.log(file);
                                    statusTableTemplateId = file.id;
                                }
                            }
                        }
                        catch (e_1_1) { e_1 = { error: e_1_1 }; }
                        finally {
                            try {
                                if (_b && !_b.done && (_c = _a.return)) _c.call(_a);
                            }
                            finally { if (e_1) throw e_1.error; }
                        }
                        if (!statusTableTemplateId) {
                            httpRes = {
                                status: 1,
                                data: null,
                                msg: "この種類の状況表テンプレートファイルが存在しないので、アップロードしてください"
                            };
                            return [2, res.json(httpRes)];
                        }
                        this.gennaretaStatusExcel(acc.Id, statusTableTemplateId, statustableSpaceId, "\u72B6\u6CC1\u8868" + statustableConfig.fileSuffix, statustableConfig, next);
                        return [2, res.json(httpRes)];
                    case 5:
                        error_1 = _d.sent();
                        console.error(error_1);
                        next(error_1);
                        httpRes.status = 2;
                        httpRes.data = error_1;
                        httpRes.msg = "システムエラーが発生しました。";
                        return [2, res.json(httpRes)];
                    case 6: return [2];
                }
            });
        }); };
        this.gennaretaStatusExcel = function (pjcode, statusTableTemplateId, statustableSpaceId, fileName, statustableConfig, next) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var fileStream, error, sfdcService, fileBuffer, uploadRes;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4, this.flkDriveService.downloadContents(statusTableTemplateId)];
                    case 1:
                        fileStream = _a.sent();
                        return [4, this.excelService.getStatustable(fileStream, pjcode, statustableConfig).catch(function (err) {
                                error = err;
                                next(error);
                            })];
                    case 2:
                        _a.sent();
                        sfdcService = new sfdc_service_1.SfdcService();
                        if (error) {
                            sfdcService.sendChatter("\n        " + fileName + "\u304C\u4F5C\u6210\u5931\u6557\u3057\u307E\u3057\u305F\u3002\u30B7\u30B9\u30C6\u30E0\u7BA1\u7406\u8005\u306B\u3054\u9023\u7D61\u4E0B\u3055\u3044\u3002\n        \u30A8\u30E9\u30FC\u30E1\u30C3\u30BB\u30FC\u30B8:\n        " + error + "\n        \u30D7\u30ED\u30B8\u30A7\u30AF\u30C8:\n      ", [statustableConfig.userId], sfdcService.getDomainUrl() + "/lightning/r/Account/" + pjcode + "/view");
                            return [2];
                        }
                        return [4, this.excelService.getFileBuffer()];
                    case 3:
                        fileBuffer = _a.sent();
                        return [4, this.flkDriveService.uploadFile2(statustableSpaceId, statustableConfig.projectName + fileName, fileBuffer)];
                    case 4:
                        uploadRes = _a.sent();
                        sfdcService.sendChatter("\n      " + fileName + "\u304C\u4F5C\u6210\u3067\u304D\u307E\u3057\u305F\u3002\u30D7\u30ED\u30B8\u30A7\u30AF\u30C8\u95A2\u9023\u3059\u308BFleekDrive\u5185\u306B\u30D5\u30A1\u30A4\u30EB\u3092\u3054\u78BA\u8A8D\u304F\u3060\u3055\u3044\u3002\n      \u30D5\u30A1\u30A4\u30EB\u540D: " + uploadRes.fileName + "\n      \u30D7\u30ED\u30B8\u30A7\u30AF\u30C8:\n    ", [statustableConfig.userId], sfdcService.getDomainUrl() + "/lightning/r/Account/" + pjcode + "/view");
                        return [2];
                }
            });
        }); };
        this.getWeeklyReport = function (req, res, next) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var httpRes, pjcode, weeklyRptConfig, acc, weeklySpaceId, error_2;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        httpRes = { status: 0, data: null, msg: "週報ファイル作成中" };
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 5]);
                        pjcode = req.params.pjcode;
                        weeklyRptConfig = req.body;
                        return [4, this.getAccount(weeklyRptConfig.projectId, pjcode)];
                    case 2:
                        acc = _a.sent();
                        if (!acc) {
                            httpRes = { status: 1, data: null, msg: "プロジェクトが存在しません" };
                            return [2, res.json(httpRes)];
                        }
                        weeklyRptConfig.projectName = acc.Name;
                        return [4, this.getSubSpaceId(acc.Id, "週報")];
                    case 3:
                        weeklySpaceId = _a.sent();
                        if (!weeklySpaceId) {
                            httpRes = { status: 1, data: null, msg: "FleekDriveに「週報」フォルダがありません。" };
                            return [2, res.json(httpRes)];
                        }
                        this.gennaretaWeeklyReport(weeklySpaceId, "週報", weeklyRptConfig, next);
                        return [2, res.json(httpRes)];
                    case 4:
                        error_2 = _a.sent();
                        console.error(error_2);
                        next(error_2);
                        return [3, 5];
                    case 5: return [2];
                }
            });
        }); };
        this.getAnalysisReport = function (req, res, next) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var httpRes, pjcode, analysisRptConfig, acc, analysisSpaceId, error_3;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        httpRes = { status: 0, data: null, msg: "分析表ファイル作成中" };
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 5]);
                        pjcode = req.params.pjcode;
                        analysisRptConfig = req.body;
                        console.log(analysisRptConfig);
                        return [4, this.getAccount(analysisRptConfig.projectId, pjcode)];
                    case 2:
                        acc = _a.sent();
                        if (!acc) {
                            httpRes = { status: 1, data: null, msg: "プロジェクトが存在しません" };
                            return [2, res.json(httpRes)];
                        }
                        analysisRptConfig.projectName = acc.Name;
                        return [4, this.getSubSpaceId(acc.Id, "分析表")];
                    case 3:
                        analysisSpaceId = _a.sent();
                        if (!analysisSpaceId) {
                            httpRes = { status: 1, data: null, msg: "FleekDriveに「分析表」フォルダがありません。" };
                            return [2, res.json(httpRes)];
                        }
                        this.gennaretaAnalysReport(analysisSpaceId, "分析表", analysisRptConfig, next);
                        return [2, res.json(httpRes)];
                    case 4:
                        error_3 = _a.sent();
                        console.error(error_3);
                        next(error_3);
                        return [2, res.json(error_3)];
                    case 5: return [2];
                }
            });
        }); };
        this.gennaretaWeeklyReport = function (weeklySpaceId, fileName, weeklyRptConfig, next) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var wps, error, sfdcService, fileBuffer, uploadRes;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        wps = new weeklyReport_service_1.WeeklyReportService();
                        return [4, wps.gennaretaWeeklyReport(weeklyRptConfig).catch(function (err) {
                                error = err;
                                console.error(err);
                                next(err);
                            })];
                    case 1:
                        _a.sent();
                        sfdcService = new sfdc_service_1.SfdcService();
                        if (error) {
                            sfdcService.sendChatter("\n        " + fileName + "\u304C\u4F5C\u6210\u5931\u6557\u3057\u307E\u3057\u305F\u3002\u30B7\u30B9\u30C6\u30E0\u7BA1\u7406\u8005\u306B\u3054\u9023\u7D61\u4E0B\u3055\u3044\u3002\n        \u30A8\u30E9\u30FC\u30E1\u30C3\u30BB\u30FC\u30B8:\n        " + error + "\n        \u30D7\u30ED\u30B8\u30A7\u30AF\u30C8:\n      ", [weeklyRptConfig.userId], sfdcService.getDomainUrl() + "/lightning/r/Account/" + weeklyRptConfig.projectId + "/view");
                            return [2];
                        }
                        return [4, wps.getFileBuffer()];
                    case 2:
                        fileBuffer = _a.sent();
                        return [4, this.flkDriveService.uploadFile2(weeklySpaceId, weeklyRptConfig.projectName + fileName, fileBuffer)];
                    case 3:
                        uploadRes = _a.sent();
                        sfdcService.sendChatter("\n      " + fileName + "\u304C\u4F5C\u6210\u3067\u304D\u307E\u3057\u305F\u3002\u30D7\u30ED\u30B8\u30A7\u30AF\u30C8\u95A2\u9023\u3059\u308BFleekDrive\u5185\u306B\u30D5\u30A1\u30A4\u30EB\u3092\u3054\u78BA\u8A8D\u304F\u3060\u3055\u3044\u3002\n      \u30D5\u30A1\u30A4\u30EB\u540D: " + uploadRes.fileName + "\n      \u30D7\u30ED\u30B8\u30A7\u30AF\u30C8:\n    ", [weeklyRptConfig.userId], sfdcService.getDomainUrl() + "/lightning/r/Account/" + weeklyRptConfig.projectId + "/view");
                        return [2];
                }
            });
        }); };
        this.gennaretaAnalysReport = function (analysisSpaceId, fileName, analysisRptConfig, next) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var ars, error, sfdcService, fileBuffer, uploadRes;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        ars = new analysisReport_service_1.AnalysReportService(analysisRptConfig, next);
                        return [4, ars.init().catch(function (err) {
                                error = err;
                                console.error(err);
                                next(err);
                            })];
                    case 1:
                        _a.sent();
                        sfdcService = new sfdc_service_1.SfdcService();
                        if (error) {
                            sfdcService.sendChatter("\n        " + fileName + "\u304C\u4F5C\u6210\u5931\u6557\u3057\u307E\u3057\u305F\u3002\u30B7\u30B9\u30C6\u30E0\u7BA1\u7406\u8005\u306B\u3054\u9023\u7D61\u4E0B\u3055\u3044\u3002\n        \u30A8\u30E9\u30FC\u30E1\u30C3\u30BB\u30FC\u30B8:\n        " + error + "\n        \u30D7\u30ED\u30B8\u30A7\u30AF\u30C8:\n      ", [analysisRptConfig.userId], sfdcService.getDomainUrl() + "/lightning/r/Account/" + analysisRptConfig.projectId + "/view");
                            return [2];
                        }
                        return [4, ars.getFileBuffer()];
                    case 2:
                        fileBuffer = _a.sent();
                        return [4, this.flkDriveService.uploadFile2(analysisSpaceId, analysisRptConfig.projectName + fileName, fileBuffer)];
                    case 3:
                        uploadRes = _a.sent();
                        sfdcService.sendChatter("\n      " + fileName + "\u304C\u4F5C\u6210\u3067\u304D\u307E\u3057\u305F\u3002\u30D7\u30ED\u30B8\u30A7\u30AF\u30C8\u95A2\u9023\u3059\u308BFleekDrive\u5185\u306B\u30D5\u30A1\u30A4\u30EB\u3092\u3054\u78BA\u8A8D\u304F\u3060\u3055\u3044\u3002\n      \u30D5\u30A1\u30A4\u30EB\u540D: " + uploadRes.fileName + "\n      \u30D7\u30ED\u30B8\u30A7\u30AF\u30C8:\n    ", [analysisRptConfig.userId], sfdcService.getDomainUrl() + "/lightning/r/Account/" + analysisRptConfig.projectId + "/view");
                        return [2];
                }
            });
        }); };
        this.getMessage = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var flkDriveService, result, spaceId, filesRes, statusTableTemplateId, _a, _b, file;
            var e_2, _c;
            return tslib_1.__generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        flkDriveService = new fleekdrive_service_1.FleekDriveService();
                        return [4, flkDriveService.authentication()];
                    case 1:
                        _d.sent();
                        return [4, flkDriveService.getRelatedListBySfdcId(undefined)];
                    case 2:
                        result = _d.sent();
                        spaceId = result.rows[0].id;
                        console.log(spaceId);
                        return [4, flkDriveService.getContentsList(spaceId)];
                    case 3:
                        filesRes = _d.sent();
                        console.log(filesRes);
                        try {
                            for (_a = tslib_1.__values(filesRes.rows), _b = _a.next(); !_b.done; _b = _a.next()) {
                                file = _b.value;
                                if (file.name === "況表テンプレート.xlsx") {
                                    statusTableTemplateId = file.id;
                                }
                            }
                        }
                        catch (e_2_1) { e_2 = { error: e_2_1 }; }
                        finally {
                            try {
                                if (_b && !_b.done && (_c = _a.return)) _c.call(_a);
                            }
                            finally { if (e_2) throw e_2.error; }
                        }
                        return [2, res.send("ok")];
                }
            });
        }); };
        this.uploadFiles = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                return [2, null];
            });
        }); };
        this.router = express_1.Router();
        this.intializeRoutes();
    }
    FilesController.prototype.intializeRoutes = function () {
        this.router.post("/api/imagecheck", this.upload.array("file", FilesController.MAX_FILES_COUNT), this.uploadFiles);
        this.router.get("/api/index", this.getMessage);
        this.router.get("/api/excel", this.getExcel);
        this.router.get("/api/test", this.getExcel2);
        this.router.post("/api/statustable/:pjcode", this.getStatustable);
        this.router.post("/api/weeklyreport/:pjcode", this.getWeeklyReport);
        this.router.post("/api/analysisreport/:pjcode", this.getAnalysisReport);
    };
    FilesController.prototype.getSubSpaceId = function (pjId, folderName) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var result, spaceId, spaceList, subSpaceId, _a, _b, space;
            var e_3, _c;
            return tslib_1.__generator(this, function (_d) {
                switch (_d.label) {
                    case 0: return [4, this.flkDriveService.authentication()];
                    case 1:
                        _d.sent();
                        return [4, this.flkDriveService.getRelatedListBySfdcId(pjId)];
                    case 2:
                        result = _d.sent();
                        if (!result || result.records === 0) {
                            return [2, null];
                        }
                        spaceId = result.rows[0].id;
                        return [4, this.flkDriveService.getSpaceContentsList(spaceId, "")];
                    case 3:
                        spaceList = _d.sent();
                        try {
                            for (_a = tslib_1.__values(spaceList.rows), _b = _a.next(); !_b.done; _b = _a.next()) {
                                space = _b.value;
                                if (space.isSpace && space.simpleName === folderName) {
                                    subSpaceId = space.id;
                                }
                            }
                        }
                        catch (e_3_1) { e_3 = { error: e_3_1 }; }
                        finally {
                            try {
                                if (_b && !_b.done && (_c = _a.return)) _c.call(_a);
                            }
                            finally { if (e_3) throw e_3.error; }
                        }
                        return [2, subSpaceId];
                }
            });
        });
    };
    FilesController.prototype.getAccount = function (projectId, pjcode) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var sfdcService, acc, sfdcResult, sfdcResult;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        sfdcService = new sfdc_service_1.SfdcService();
                        if (!!projectId) return [3, 2];
                        return [4, sfdcService.query("SELECT Id,Name FROM Account WHERE ProjectCode__c = '" + pjcode + "'")];
                    case 1:
                        sfdcResult = _a.sent();
                        if (sfdcResult && sfdcResult.totalSize > 0) {
                            acc = sfdcResult.records[0];
                        }
                        return [3, 4];
                    case 2: return [4, sfdcService.query("SELECT Id,Name FROM Account WHERE Id = '" + projectId + "'")];
                    case 3:
                        sfdcResult = _a.sent();
                        if (sfdcResult && sfdcResult.totalSize > 0) {
                            acc = sfdcResult.records[0];
                        }
                        _a.label = 4;
                    case 4: return [2, acc];
                }
            });
        });
    };
    FilesController.MAX_FILES_COUNT = 10;
    return FilesController;
}());
exports.FilesController = FilesController;
//# sourceMappingURL=files.controller.js.map