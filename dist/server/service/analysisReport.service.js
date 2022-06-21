"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.AnalysReportService = void 0;
var tslib_1 = require("tslib");
var exceljs_1 = tslib_1.__importDefault(require("exceljs"));
var sfdc_service_1 = require("./sfdc.service");
var csvtojson_1 = tslib_1.__importDefault(require("csvtojson"));
var tools_service_1 = require("./tools.service");
var AnalysisPrefecturesInfo = (function () {
    function AnalysisPrefecturesInfo() {
        this.name = "";
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    return AnalysisPrefecturesInfo;
}());
var PriorityAreaInfo = (function () {
    function PriorityAreaInfo() {
        this.prefectures = "";
        this.city = "";
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    PriorityAreaInfo.AreaTotalInfo = {
        sum1Total: 0,
        sum2Total: 0
    };
    return PriorityAreaInfo;
}());
var AgeInfo = (function () {
    function AgeInfo() {
        this.agePeriod = "";
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    AgeInfo.getAgePeriod = function (age) {
        return ~~(age / 5) * 5 + "歳";
    };
    AgeInfo.ageTotalInfo = {
        sum1Total: 0,
        sum2Total: 0,
        total: 0
    };
    return AgeInfo;
}());
var FamiliesInfo = (function () {
    function FamiliesInfo() {
        this.peopleNumber = 0;
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    FamiliesInfo.TotalInfo = {
        sum1Total: 0,
        sum2Total: 0,
        total: 0
    };
    return FamiliesInfo;
}());
var AnnualIncomeInfo = (function () {
    function AnnualIncomeInfo() {
        this.aggregateName = "";
        this.amountStart = 0;
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    AnnualIncomeInfo.TotalInfo = {
        sum1Total: 0,
        sum2Total: 0,
        total: 0
    };
    return AnnualIncomeInfo;
}());
var OwnResourcesInfo = (function () {
    function OwnResourcesInfo() {
        this.aggregateName = "";
        this.amountStart = 0;
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    OwnResourcesInfo.TotalInfo = {
        sum1Total: 0,
        sum2Total: 0,
        total: 0
    };
    return OwnResourcesInfo;
}());
var BudgetInfo = (function () {
    function BudgetInfo() {
        this.aggregateName = "";
        this.amountStart = 0;
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    BudgetInfo.TotalInfo = {
        sum1Total: 0,
        sum2Total: 0,
        total: 0
    };
    return BudgetInfo;
}());
var DesiredFloorPlan = (function () {
    function DesiredFloorPlan() {
        this.name = "";
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    return DesiredFloorPlan;
}());
var DesiredFloorPlanInfo = (function () {
    function DesiredFloorPlanInfo() {
        this.plan1TotalInfo = {
            sum1Total: 0,
            sum2Total: 0,
            total: 0
        };
        this.plan2TotalInfo = {
            sum1Total: 0,
            sum2Total: 0,
            total: 0
        };
        this.plan1 = new Map();
        this.plan2 = new Map();
    }
    return DesiredFloorPlanInfo;
}());
var DesiredAreaInfo = (function () {
    function DesiredAreaInfo() {
        this.aggregateName = "";
        this.numStart = 0;
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    DesiredAreaInfo.TotalInfo = {
        sum1Total: 0,
        sum2Total: 0,
        total: 0
    };
    return DesiredAreaInfo;
}());
var LivingStyleInfo = (function () {
    function LivingStyleInfo() {
        this.name = "";
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    LivingStyleInfo.TotalInfo = {
        sum1Total: 0,
        sum2Total: 0,
        total: 0
    };
    return LivingStyleInfo;
}());
var ReplacementScheduleInfo = (function () {
    function ReplacementScheduleInfo() {
        this.name = "";
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    ReplacementScheduleInfo.TotalInfo = {
        sum1Total: 0,
        sum2Total: 0,
        total: 0
    };
    return ReplacementScheduleInfo;
}());
var ReplacementPropertyInfo = (function () {
    function ReplacementPropertyInfo() {
        this.name = "";
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    ReplacementPropertyInfo.TotalInfo = {
        sum1Total: 0,
        sum2Total: 0,
        total: 0
    };
    return ReplacementPropertyInfo;
}());
var ProfessionInfo = (function () {
    function ProfessionInfo() {
        this.name = "";
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    ProfessionInfo.TotalInfo = {
        sum1Total: 0,
        sum2Total: 0,
        total: 0
    };
    return ProfessionInfo;
}());
var WorkPlaceInfo = (function () {
    function WorkPlaceInfo() {
        this.name = "";
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    WorkPlaceInfo.TotalInfo = {
        sum1Total: 0,
        sum2Total: 0,
        total: 0
    };
    return WorkPlaceInfo;
}());
var ParkingRequestInfo = (function () {
    function ParkingRequestInfo() {
        this.name = "";
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    ParkingRequestInfo.TotalInfo = {
        sum1Total: 0,
        sum2Total: 0,
        total: 0
    };
    return ParkingRequestInfo;
}());
var ParkingNumberInfo = (function () {
    function ParkingNumberInfo() {
        this.name = "";
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    ParkingNumberInfo.TotalInfo = {
        sum1Total: 0,
        sum2Total: 0,
        total: 0
    };
    return ParkingNumberInfo;
}());
var MediumInfo = (function () {
    function MediumInfo() {
        this.name = "";
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    return MediumInfo;
}());
var filedMappingMedia2 = {
    MediumNewspaperFlyers__c: "新聞折込チラシ",
    MediumHouseInformationMagazine__c: "住宅情報誌",
    MediumWorkPlaceExaminationMedium__c: "ご勤務先",
    MediumAcquaintancesIntroduction__c: "知人等のご紹介",
    MediumSellerMembershipOrganization__c: "売主の会員組織",
    MediumTrafficAdvertising__c: "交通広告",
    MediumMassMedia__c: "マスメディア",
    MediumArticlesLocalSignboards__c: "投函物・現地看板等"
};
var AnswerInfo = (function () {
    function AnswerInfo() {
        this.name = "";
        this.itemList = new Map();
        this.countNum1 = 0;
        this.countNum2 = 0;
        this.total = 0;
    }
    return AnswerInfo;
}());
var AnalysReportService = (function () {
    function AnalysReportService(analysisRptConfig, next) {
        this.PageInfo = {
            pageSize: 60,
            lineStart: 9,
            indexNum: 12,
            colNum: 0,
            get colIndex() {
                return this.colNum * 9 + 1;
            },
            get hasRow() {
                if (this.pageSize === 60) {
                    return this.pageSize - this.indexNum;
                }
                else {
                    return this.pageSize - ((this.indexNum + 2) % this.pageSize);
                }
            },
            get rowIndex() {
                return this.indexNum;
            },
            runIndexByNum: function (loopNum) {
                for (var index = 0; index < loopNum; index++) {
                    this.runIndex();
                }
            },
            runIndex: function () {
                if ((this.pageSize === 60 && this.indexNum % this.pageSize === 0) ||
                    (this.pageSize !== 60 &&
                        this.indexNum > AnalysReportService.PAGE_SIZE &&
                        (this.indexNum + 2) % this.pageSize === 0)) {
                    this.colNum++;
                    if (this.colNum === 3) {
                        this.colNum = 0;
                        this.indexNum++;
                        this.lineStart = this.indexNum;
                        if (this.pageSize === 60) {
                            this.pageSize = AnalysReportService.PAGE_SIZE;
                        }
                    }
                    else {
                        this.indexNum = this.lineStart;
                    }
                }
                this.indexNum++;
            }
        };
        this.MediumPageInfo = {
            rowIndex: 11,
            colNum: 0,
            get colIndex() {
                return this.colNum * 13 + 1;
            },
            init: function () {
                this.rowIndex = 11;
                this.colNum = 0;
            },
            runIndex: function (num) {
                if (num) {
                    this.rowIndex += num;
                }
                else {
                    this.rowIndex++;
                }
            }
        };
        this.OtherPageInfo = {
            pageSize: 64,
            lineStart: 7,
            indexNum: 7,
            colNum: 0,
            get colIndex() {
                return this.colNum * 9 + 1;
            },
            get hasRow() {
                return this.pageSize - (this.indexNum % this.pageSize);
            },
            get rowIndex() {
                return this.indexNum;
            },
            runIndexByNum: function (loopNum) {
                for (var index = 0; index < loopNum; index++) {
                    this.runIndex();
                }
            },
            runIndex: function () {
                if (this.indexNum % this.pageSize === 0) {
                    this.colNum++;
                    if (this.colNum === 3) {
                        this.colNum = 0;
                        this.indexNum++;
                        this.lineStart = this.indexNum;
                    }
                    else {
                        this.indexNum = this.lineStart;
                    }
                }
                this.indexNum++;
            }
        };
        this.preTotalInfo = {
            recordsTotal: 0,
            sum1Total: 0,
            sum2Total: 0
        };
        this.md1TotalInfo = {
            sum1Total: 0,
            sum2Total: 0,
            Total: 0
        };
        this.md2TotalInfo = {
            sum1Total: 0,
            sum2Total: 0,
            Total: 0,
            noneSum1: 0,
            noneSum2: 0,
            noneTotal: 0
        };
        this.md3TotalInfo = {
            sum1Total: 0,
            sum2Total: 0,
            Total: 0,
            noneSum1: 0,
            noneSum2: 0,
            noneTotal: 0
        };
        this.workbook = new exceljs_1.default.Workbook();
        this.sheetLsit = [];
        this.sfdcService = new sfdc_service_1.SfdcService();
        this.analysisRptConfig = analysisRptConfig;
        this.comPareDate = {
            period1Start: tools_service_1.Tools.strToDate(analysisRptConfig.period1Start),
            period1End: tools_service_1.Tools.strToDate(analysisRptConfig.period1End),
            period2Start: tools_service_1.Tools.strToDate(analysisRptConfig.period2Start),
            period2End: tools_service_1.Tools.strToDate(analysisRptConfig.period2End)
        };
        this.next = next;
    }
    AnalysReportService.prototype.init = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _a = this;
                        return [4, this.workbook.xlsx.readFile("server/template/分析表テンプレート.xlsx")];
                    case 1:
                        _a.workbook = _b.sent();
                        this.sheetLsit = this.workbook.worksheets;
                        return [4, this.setAnasislyTable().catch(function (err) {
                                throw err;
                            })];
                    case 2:
                        _b.sent();
                        return [2];
                }
            });
        });
    };
    AnalysReportService.prototype.getFileBuffer = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4, this.workbook.xlsx.writeBuffer()];
                    case 1: return [2, _a.sent()];
                }
            });
        });
    };
    AnalysReportService.prototype.setAnasislyTable = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var anasislyTableWorkSheet, mediaWorkSheet, otherSheet, formatSheet, promiseJobList;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        anasislyTableWorkSheet = this.getWorkSheetByName("\u5E33\u7968\u30EC\u30A4\u30A2\u30A6\u30C8\uFF08XX\u5206\u6790\u8868\uFF09");
                        mediaWorkSheet = this.getWorkSheetByName("\u5E33\u7968\u30EC\u30A4\u30A2\u30A6\u30C8\uFF08XX\u7533\u544A\u5A92\u4F53\u5206\u6790\uFF09");
                        otherSheet = this.getWorkSheetByName("\u5E33\u7968\u30EC\u30A4\u30A2\u30A6\u30C8\uFF08XX\u305D\u306E\u4ED6\u5206\u6790\u8868\uFF09");
                        formatSheet = this.getWorkSheetByName("XX分析表");
                        if (!formatSheet)
                            return [2];
                        promiseJobList = [];
                        promiseJobList.push(this.setTitle([anasislyTableWorkSheet, mediaWorkSheet, otherSheet]));
                        promiseJobList.push(this.setPrefectures());
                        if (anasislyTableWorkSheet) {
                            this.setPriorityArea();
                            this.optAgeMap = new Map();
                            this.optFamiliesMap = new Map();
                            this.optAnnualIncomeMap = new Map();
                            this.optOwnResources = new Map();
                            this.optBudgetMap = new Map();
                            this.desiredFloorPlanInfo = new DesiredFloorPlanInfo();
                            this.optDesiredAreaMap = new Map();
                            this.optLivingStyleMap = new Map();
                            this.optReplaceMap = new Map();
                            this.optReplacePropertyMap = new Map();
                            this.optProfessionMap = new Map();
                            this.optWorkPlaceMap = new Map();
                            this.optParkingRequestMap = new Map();
                            this.optParkingNumberMap = new Map();
                        }
                        if (mediaWorkSheet) {
                            this.optMediumMap1 = new Map();
                            this.optMediumMap2 = new Map();
                            this.optMediumMap3 = new Map();
                        }
                        if (otherSheet) {
                            promiseJobList.push(this.setOtherAnalysis());
                        }
                        return [4, Promise.all(promiseJobList).catch(function (err) {
                                throw err;
                            })];
                    case 1:
                        _a.sent();
                        return [4, this.sumDataByBulkRecord()];
                    case 2:
                        _a.sent();
                        promiseJobList = [];
                        if (anasislyTableWorkSheet) {
                            this.outPutPrefectures(anasislyTableWorkSheet, formatSheet);
                            this.outPutPriorityArea(anasislyTableWorkSheet, formatSheet);
                            this.outPutAgeInfo(anasislyTableWorkSheet, formatSheet);
                            this.outPutFamilieInfo(anasislyTableWorkSheet, formatSheet);
                            this.outPutAnnualIncome(anasislyTableWorkSheet, formatSheet);
                            this.outPutOwnResources(anasislyTableWorkSheet, formatSheet);
                            this.outPutBudget(anasislyTableWorkSheet, formatSheet);
                            this.outDesiredFloorPlan(anasislyTableWorkSheet, formatSheet, "plan1");
                            this.outDesiredFloorPlan(anasislyTableWorkSheet, formatSheet, "plan2");
                            this.outputDesiredArea(anasislyTableWorkSheet, formatSheet);
                            this.outPutLivingStyle(anasislyTableWorkSheet, formatSheet);
                            this.outPutReplacement(anasislyTableWorkSheet, formatSheet);
                            this.outPutReplacementProperty(anasislyTableWorkSheet, formatSheet);
                            this.outPutProfession(anasislyTableWorkSheet, formatSheet);
                            this.outPutWorkPlace(anasislyTableWorkSheet, formatSheet);
                            this.outPutParkingRequest(anasislyTableWorkSheet, formatSheet);
                            this.outPutParkingNumber(anasislyTableWorkSheet, formatSheet);
                        }
                        if (mediaWorkSheet) {
                            this.outPutMedium1(mediaWorkSheet, formatSheet);
                            this.outPutMedium2(mediaWorkSheet, formatSheet);
                            this.outPutMedium3(mediaWorkSheet, formatSheet);
                        }
                        if (otherSheet) {
                            this.outPutOther(otherSheet, formatSheet);
                        }
                        return [4, Promise.all(promiseJobList).catch(function (err) {
                                throw err;
                            })];
                    case 3:
                        _a.sent();
                        this.workbook.removeWorksheetEx(formatSheet);
                        return [2];
                }
            });
        });
    };
    AnalysReportService.prototype.setTitle = function (workSheets) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var pjRes, p1, p2, p3, workSheets_1, workSheets_1_1, workSheet;
            var e_1, _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0: return [4, this.sfdcService.query("SELECT Id,Name FROM Account WHERE Id = '" + this.analysisRptConfig.projectId + "'")];
                    case 1:
                        pjRes = _b.sent();
                        if (pjRes && pjRes.totalSize > 0) {
                            this.analysisRptConfig.projectName = pjRes.records[0].Name;
                        }
                        return [4, this.sfdcService.query("SELECT count(Id) countNum FROM OpportunityHistory__c WHERE Opportunity__r.AccountId = '" + this.analysisRptConfig.projectId + "' AND ResponseDate__c >= " + this.analysisRptConfig.period1Start + " AND ResponseDate__c <= " + this.analysisRptConfig.period1End + " AND StageName__c = '" + this.analysisRptConfig.stageName + "'")];
                    case 2:
                        p1 = _b.sent();
                        return [4, this.sfdcService.query("SELECT count(Id) countNum FROM OpportunityHistory__c WHERE Opportunity__r.AccountId = '" + this.analysisRptConfig.projectId + "' AND ResponseDate__c >= " + this.analysisRptConfig.period2Start + " AND ResponseDate__c <= " + this.analysisRptConfig.period2End + " AND StageName__c = '" + this.analysisRptConfig.stageName + "'")];
                    case 3:
                        p2 = _b.sent();
                        return [4, this.sfdcService.query("SELECT count(Id) countNum FROM OpportunityHistory__c WHERE Opportunity__r.AccountId = '" + this.analysisRptConfig.projectId + "' AND StageName__c = '" + this.analysisRptConfig.stageName + "'")];
                    case 4:
                        p3 = _b.sent();
                        try {
                            for (workSheets_1 = tslib_1.__values(workSheets), workSheets_1_1 = workSheets_1.next(); !workSheets_1_1.done; workSheets_1_1 = workSheets_1.next()) {
                                workSheet = workSheets_1_1.value;
                                if (!workSheet)
                                    continue;
                                workSheet.getCell("A1").value = "\u7269\u4EF6\u540D\uFF1A" + this.analysisRptConfig.projectName;
                                workSheet.getCell("A2").value = "\u671F\u95931\uFF1A" + this.analysisRptConfig.period1Start.replace(/-/g, "/") + "\uFF5E" + this.analysisRptConfig.period1End.replace(/-/g, "/");
                                workSheet.getCell("A3").value = "\u671F\u95932\uFF1A" + this.analysisRptConfig.period2Start.replace(/-/g, "/") + "\uFF5E" + this.analysisRptConfig.period2End.replace(/-/g, "/");
                                if (p1) {
                                    workSheet.getCell("J2").value = p1.records[0].countNum;
                                }
                                if (p2) {
                                    workSheet.getCell("J3").value = p2.records[0].countNum;
                                }
                                if (p3) {
                                    workSheet.getCell("J4").value = p3.records[0].countNum;
                                }
                            }
                        }
                        catch (e_1_1) { e_1 = { error: e_1_1 }; }
                        finally {
                            try {
                                if (workSheets_1_1 && !workSheets_1_1.done && (_a = workSheets_1.return)) _a.call(workSheets_1);
                            }
                            finally { if (e_1) throw e_1.error; }
                        }
                        return [2];
                }
            });
        });
    };
    AnalysReportService.prototype.setPrefectures = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var optMap, showCityList, _a, _b, pre, asp_1, adRes, _c, _d, address, asp_2, city, asp;
            var e_2, _e, e_3, _f;
            return tslib_1.__generator(this, function (_g) {
                switch (_g.label) {
                    case 0:
                        optMap = new Map();
                        this.optMap = optMap;
                        showCityList = [];
                        try {
                            for (_a = tslib_1.__values(this.analysisRptConfig.prefectureList), _b = _a.next(); !_b.done; _b = _a.next()) {
                                pre = _b.value;
                                asp_1 = new AnalysisPrefecturesInfo();
                                asp_1.name = pre.Prefectures__c;
                                asp_1.analysisDisplayClassification = pre.AnalysisDisplayClassification__c;
                                if (asp_1.analysisDisplayClassification === "市区町村を表示する") {
                                    showCityList.push(asp_1.name);
                                    asp_1.cityList = new Map();
                                }
                                optMap.set(asp_1.name, asp_1);
                            }
                        }
                        catch (e_2_1) { e_2 = { error: e_2_1 }; }
                        finally {
                            try {
                                if (_b && !_b.done && (_e = _a.return)) _e.call(_a);
                            }
                            finally { if (e_2) throw e_2.error; }
                        }
                        if (showCityList.length === 0)
                            return [2];
                        return [4, this.sfdcService.query("SELECT Prefectures__c,City__c FROM AddressMst__c WHERE Prefectures__c IN " + sfdc_service_1.SfdcService.getInSql(optMap.keys()) + " GROUP BY Prefectures__c,City__c ")];
                    case 1:
                        adRes = _g.sent();
                        if (!adRes || adRes.totalSize === 0)
                            return [2];
                        try {
                            for (_c = tslib_1.__values(adRes.records), _d = _c.next(); !_d.done; _d = _c.next()) {
                                address = _d.value;
                                asp_2 = optMap.get(address.Prefectures__c);
                                if (!asp_2) {
                                    asp_2 = new AnalysisPrefecturesInfo();
                                    asp_2.name = address.Prefectures__c;
                                    asp_2.analysisDisplayClassification = "合計行を表示する";
                                    optMap.set(asp_2.name, asp_2);
                                }
                                if (!asp_2.cityList || !address.City__c)
                                    continue;
                                city = new AnalysisPrefecturesInfo();
                                city.name = address.City__c;
                                asp_2.cityList.set(city.name, city);
                            }
                        }
                        catch (e_3_1) { e_3 = { error: e_3_1 }; }
                        finally {
                            try {
                                if (_d && !_d.done && (_f = _c.return)) _f.call(_c);
                            }
                            finally { if (e_3) throw e_3.error; }
                        }
                        asp = new AnalysisPrefecturesInfo();
                        asp.name = "海外";
                        asp.analysisDisplayClassification = "合計行を表示する";
                        optMap.set(asp.name, asp);
                        return [2];
                }
            });
        });
    };
    AnalysReportService.prototype.sumDataByBulkRecord = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var recordStream, csvToJsonParser, readStream;
            var _this = this;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4, this.sfdcService.bulkQuery("SELECT Id,Prefectures__c,City__c,Towns__c,WeeklyReportAnalysisCity__c,WeeklyReportAnalysisChome__c,ResponseDate__c,OverSeaFlag__c,ResponsingAge__c,NumberPeoplePlanMovein__c,GenerationMainAnnualIncome__c,DesiredFloorPlan1__c,DesiredFloorPlan2__c,OwnResources__c,Budget__c,DesiredArea__c,LivingStyle__c,ReplacementSchedule__c,ReplacementProperty__c,Profession__c,WorkPlacePrefectures__c,WorkPlaceCity__c,ParkingRequest__c,ParkingNumber__c,MediumInternet__c,MajorItems__c,SubItem__c,MediumNewspaperFlyers__c,MediumHouseInformationMagazine__c,MediumWorkPlaceExaminationMedium__c,MediumAcquaintancesIntroduction__c,MediumSellerMembershipOrganization__c,MediumTrafficAdvertising__c,MediumMassMedia__c,MediumArticlesLocalSignboards__c FROM OpportunityHistory__c WHERE Opportunity__r.AccountId = '" + this.analysisRptConfig.projectId + "' AND StageName__c = '" + this.analysisRptConfig.stageName + "' AND (Prefectures__c != null OR OverSeaFlag__c = '\u6D77\u5916')")];
                    case 1:
                        recordStream = _a.sent();
                        csvToJsonParser = csvtojson_1.default({ checkType: true });
                        readStream = recordStream.stream();
                        readStream.pipe(csvToJsonParser);
                        csvToJsonParser.on("data", function (data) {
                            _this.preTotalInfo.recordsTotal++;
                            var record = JSON.parse(data.toString("utf8"));
                            _this.sumPrefecturesInfo(record, _this.optMap, _this.preTotalInfo);
                            _this.sumPriorityArea(record, _this.optAreaMap);
                            _this.sumAge(record, _this.optAgeMap);
                            _this.sumFamilies(record, _this.optFamiliesMap);
                            _this.sumAnnualIncome(record, _this.optAnnualIncomeMap);
                            _this.sumOwnResources(record, _this.optOwnResources);
                            _this.sumBudget(record, _this.optBudgetMap);
                            _this.sumDesiredFloorPlan(record, "plan1", _this.desiredFloorPlanInfo);
                            _this.sumDesiredFloorPlan(record, "plan2", _this.desiredFloorPlanInfo);
                            _this.sumDesiredArea(record, _this.optDesiredAreaMap);
                            _this.sumLivingStyle(record, _this.optLivingStyleMap);
                            _this.sumReplacement(record, _this.optReplaceMap);
                            _this.sumReplacementProperty(record, _this.optReplacePropertyMap);
                            _this.sumProfession(record, _this.optProfessionMap);
                            _this.sumWorkPlace(record, _this.optWorkPlaceMap);
                            _this.sumParkingRequest(record, _this.optParkingRequestMap);
                            _this.sumParkingNumber(record, _this.optParkingNumberMap);
                            _this.sumMedium1(record);
                            _this.sumMedium2(record);
                            _this.sumMedium3(record);
                        });
                        return [4, new Promise(function (resolve, reject) {
                                recordStream.on("error", function (err) {
                                    console.error(err);
                                    reject(new Error("Couldn't download results from Salesforce Bulk API"));
                                });
                                csvToJsonParser.on("error", function (error) {
                                    console.error(error);
                                    reject(new Error("Couldn't parse results from Salesforce Bulk API"));
                                });
                                csvToJsonParser.on("done", function () { return tslib_1.__awaiter(_this, void 0, void 0, function () {
                                    return tslib_1.__generator(this, function (_a) {
                                        resolve("ok");
                                        return [2];
                                    });
                                }); });
                            }).then(function () {
                                console.log("sucees record total:" + _this.preTotalInfo.recordsTotal);
                            })];
                    case 2:
                        _a.sent();
                        return [2];
                }
            });
        });
    };
    AnalysReportService.prototype.getPeriodByDate = function (responseDate) {
        var setNum;
        if (this.comPareDate.period1Start &&
            this.comPareDate.period1End &&
            responseDate >= this.comPareDate.period1Start &&
            responseDate <= this.comPareDate.period1End) {
            setNum = 1;
        }
        else if (this.comPareDate.period2Start &&
            this.comPareDate.period2End &&
            responseDate >= this.comPareDate.period2Start &&
            responseDate <= this.comPareDate.period2End) {
            setNum = 2;
        }
        return setNum;
    };
    AnalysReportService.prototype.sumPrefecturesInfo = function (record, optMap, preTotalInfo) {
        if (!optMap)
            return;
        var preInfo = optMap.get(record.OverSeaFlag__c ? record.OverSeaFlag__c : record.Prefectures__c);
        if (!preInfo)
            return;
        preInfo.total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            preInfo.countNum1++;
            preTotalInfo.sum1Total++;
        }
        else {
            preInfo.countNum2++;
            preTotalInfo.sum2Total++;
        }
        if (preInfo.analysisDisplayClassification === "合計行を表示する")
            return;
        if (preInfo.cityList && record.City__c) {
            var cityInfo = preInfo.cityList.get(record.City__c);
            if (cityInfo) {
                if (setNum === 1) {
                    cityInfo.countNum1++;
                }
                else {
                    cityInfo.countNum2++;
                }
            }
        }
    };
    AnalysReportService.prototype.outPutPrefectures = function (workSheet, formatSheet) {
        var e_4, _a, e_5, _b;
        var optMap = this.optMap;
        if (!optMap)
            return;
        var titleFormat = formatSheet.getCell("A6");
        var cityFormat = [];
        formatSheet.getRow(7).eachCell(function (cell) {
            cityFormat.push(cell);
        });
        var totalFormat = [];
        formatSheet.getRow(9).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        var title = "";
        try {
            for (var _c = tslib_1.__values(optMap.values()), _d = _c.next(); !_d.done; _d = _c.next()) {
                var preInfo = _d.value;
                if (title != preInfo.name) {
                    var targetCell = workSheet.getCell(this.PageInfo.rowIndex, this.PageInfo.colIndex);
                    targetCell.style = titleFormat.style;
                    targetCell.value = preInfo.name;
                    this.PageInfo.runIndex();
                    title = preInfo.name;
                }
                if (preInfo.analysisDisplayClassification === "市区町村を表示する" &&
                    preInfo.cityList &&
                    preInfo.cityList.size > 0) {
                    try {
                        for (var _e = (e_5 = void 0, tslib_1.__values(preInfo.cityList.values())), _f = _e.next(); !_f.done; _f = _e.next()) {
                            var cityInfo = _f.value;
                            var cells_1 = this.getCellsByFormat(workSheet, cityFormat);
                            cells_1[0].value = cityInfo.name;
                            cells_1[1].merge(cells_1[0]);
                            cells_1[2].value = cityInfo.countNum1;
                            cells_1[3].value = preInfo.countNum1 === 0 ? 0 : cityInfo.countNum1 / preInfo.countNum1;
                            cells_1[4].value = cityInfo.countNum2;
                            cells_1[5].value = preInfo.countNum2 === 0 ? 0 : cityInfo.countNum2 / preInfo.countNum2;
                            this.PageInfo.runIndex();
                            cells_1[6].value = cityInfo.total;
                            cells_1[7].value = preInfo.total === 0 ? 0 : cityInfo.total / preInfo.total;
                        }
                    }
                    catch (e_5_1) { e_5 = { error: e_5_1 }; }
                    finally {
                        try {
                            if (_f && !_f.done && (_b = _e.return)) _b.call(_e);
                        }
                        finally { if (e_5) throw e_5.error; }
                    }
                    this.PageInfo.runIndex();
                    this.optTotalForPrefectures(preInfo, workSheet, totalFormat);
                }
                else {
                    this.optTotalForPrefectures(preInfo, workSheet, totalFormat);
                }
            }
        }
        catch (e_4_1) { e_4 = { error: e_4_1 }; }
        finally {
            try {
                if (_d && !_d.done && (_a = _c.return)) _a.call(_c);
            }
            finally { if (e_4) throw e_4.error; }
        }
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = this.preTotalInfo.sum1Total;
        cells[3].merge(cells[2]);
        cells[4].value = this.preTotalInfo.sum2Total;
        cells[5].merge(cells[4]);
        cells[6].value = this.preTotalInfo.recordsTotal;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.sumWorkPlace = function (record, optWorkPlaceMap) {
        if (!optWorkPlaceMap)
            return;
        if (!record.OverSeaFlag__c && !record.WorkPlacePrefectures__c)
            return;
        var workPlaceName = record.OverSeaFlag__c ? record.OverSeaFlag__c : record.WorkPlacePrefectures__c;
        var preInfo = optWorkPlaceMap.get(workPlaceName);
        if (!preInfo) {
            preInfo = new WorkPlaceInfo();
            preInfo.name = workPlaceName;
            preInfo.cityList = new Map();
            optWorkPlaceMap.set(preInfo.name, preInfo);
        }
        preInfo.total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            preInfo.countNum1++;
            WorkPlaceInfo.TotalInfo.sum1Total++;
        }
        else {
            preInfo.countNum2++;
            WorkPlaceInfo.TotalInfo.sum2Total++;
        }
        if (preInfo.cityList && record.WorkPlaceCity__c) {
            var cityInfo = preInfo.cityList.get(record.WorkPlaceCity__c);
            if (!cityInfo) {
                cityInfo = new WorkPlaceInfo();
                cityInfo.name = record.WorkPlaceCity__c;
                preInfo.cityList.set(cityInfo.name, cityInfo);
            }
            if (setNum === 1) {
                cityInfo.countNum1++;
            }
            else {
                cityInfo.countNum2++;
            }
        }
    };
    AnalysReportService.prototype.outPutWorkPlace = function (workSheet, formatSheet) {
        var e_6, _a, e_7, _b;
        var optWorkPlaceMap = this.optWorkPlaceMap;
        if (!optWorkPlaceMap)
            return;
        var titleFormat = formatSheet.getCell("A6");
        var cityFormat = [];
        formatSheet.getRow(7).eachCell(function (cell) {
            cityFormat.push(cell);
        });
        var totalFormat = [];
        formatSheet.getRow(9).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        this.setPaTitle(workSheet, formatSheet, 78);
        var title = "";
        try {
            for (var _c = tslib_1.__values(optWorkPlaceMap.values()), _d = _c.next(); !_d.done; _d = _c.next()) {
                var preInfo = _d.value;
                if (title != preInfo.name) {
                    var targetCell = workSheet.getCell(this.PageInfo.rowIndex, this.PageInfo.colIndex);
                    targetCell.style = titleFormat.style;
                    targetCell.value = preInfo.name;
                    this.PageInfo.runIndex();
                    title = preInfo.name;
                }
                if (preInfo.cityList && preInfo.cityList.size > 0) {
                    try {
                        for (var _e = (e_7 = void 0, tslib_1.__values(preInfo.cityList.values())), _f = _e.next(); !_f.done; _f = _e.next()) {
                            var cityInfo = _f.value;
                            var cells_2 = this.getCellsByFormat(workSheet, cityFormat);
                            cells_2[0].value = cityInfo.name;
                            cells_2[1].merge(cells_2[0]);
                            cells_2[2].value = cityInfo.countNum1;
                            cells_2[3].value = preInfo.countNum1 === 0 ? 0 : cityInfo.countNum1 / preInfo.countNum1;
                            cells_2[4].value = cityInfo.countNum2;
                            cells_2[5].value = preInfo.countNum2 === 0 ? 0 : cityInfo.countNum2 / preInfo.countNum2;
                            this.PageInfo.runIndex();
                            cells_2[6].value = cityInfo.total;
                            cells_2[7].value = preInfo.total === 0 ? 0 : cityInfo.total / preInfo.total;
                        }
                    }
                    catch (e_7_1) { e_7 = { error: e_7_1 }; }
                    finally {
                        try {
                            if (_f && !_f.done && (_b = _e.return)) _b.call(_e);
                        }
                        finally { if (e_7) throw e_7.error; }
                    }
                    this.PageInfo.runIndex();
                    this.optTotalForPrefectures(preInfo, workSheet, totalFormat);
                }
                else {
                    this.optTotalForPrefectures(preInfo, workSheet, totalFormat);
                }
            }
        }
        catch (e_6_1) { e_6 = { error: e_6_1 }; }
        finally {
            try {
                if (_d && !_d.done && (_a = _c.return)) _a.call(_c);
            }
            finally { if (e_6) throw e_6.error; }
        }
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = this.preTotalInfo.sum1Total;
        cells[3].merge(cells[2]);
        cells[4].value = this.preTotalInfo.sum2Total;
        cells[5].merge(cells[4]);
        cells[6].value = this.preTotalInfo.recordsTotal;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.optTotalForPrefectures = function (preInfo, workSheet, totalFormat) {
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = preInfo.name + "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = preInfo.countNum1;
        cells[3].merge(cells[2]);
        cells[4].value = preInfo.countNum2;
        cells[5].merge(cells[4]);
        cells[6].value = preInfo.total;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.optTotalForMedium = function (preInfo, totalInfo, workSheet, totalFormat) {
        var cells = this.getCellsByFormat(workSheet, totalFormat, this.MediumPageInfo);
        cells[0].value = preInfo.name + " \u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].merge(cells[0]);
        cells[3].value = preInfo.countNum1;
        cells[5].value = totalInfo.sum1Total === 0 ? 0 : preInfo.countNum1 / totalInfo.sum1Total;
        cells[6].value = preInfo.countNum2;
        cells[8].value = totalInfo.sum2Total === 0 ? 0 : preInfo.countNum2 / totalInfo.sum2Total;
        cells[9].value = preInfo.total;
        cells[11].value = totalInfo.Total === 0 ? 0 : preInfo.total / totalInfo.Total;
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.optTotalForArea = function (preInfo, workSheet, totalFormat) {
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = preInfo.city + "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = preInfo.countNum1;
        cells[3].merge(cells[2]);
        cells[4].value = preInfo.countNum2;
        cells[5].merge(cells[4]);
        cells[6].value = preInfo.total;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.setPriorityArea = function () {
        var e_8, _a;
        if (!this.analysisRptConfig.priorityAreaList || this.analysisRptConfig.priorityAreaList.length === 0)
            return;
        this.optAreaMap = new Map();
        try {
            for (var _b = tslib_1.__values(this.analysisRptConfig.priorityAreaList), _c = _b.next(); !_c.done; _c = _b.next()) {
                var pArea = _c.value;
                var paInfo = new PriorityAreaInfo();
                paInfo.prefectures = pArea.Prefectures__c;
                paInfo.city = pArea.City__c;
                paInfo.town = pArea.Towns__c;
                this.optAreaMap.set(pArea.Prefectures__c + pArea.City__c, paInfo);
            }
        }
        catch (e_8_1) { e_8 = { error: e_8_1 }; }
        finally {
            try {
                if (_c && !_c.done && (_a = _b.return)) _a.call(_b);
            }
            finally { if (e_8) throw e_8.error; }
        }
    };
    AnalysReportService.prototype.sumPriorityArea = function (record, optAreaMap) {
        if (!optAreaMap || !record.Prefectures__c || !record.City__c)
            return;
        var areaInfo = optAreaMap.get(record.Prefectures__c + record.City__c);
        if (!areaInfo)
            return;
        areaInfo.total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            areaInfo.countNum1++;
            PriorityAreaInfo.AreaTotalInfo.sum1Total++;
        }
        else {
            areaInfo.countNum2++;
            PriorityAreaInfo.AreaTotalInfo.sum2Total++;
        }
        if (!record.WeeklyReportAnalysisChome__c)
            return;
        if (!areaInfo.chomeList) {
            areaInfo.chomeList = new Map();
        }
        var chomeInfo = areaInfo.chomeList.get(record.WeeklyReportAnalysisChome__c);
        if (areaInfo.town && record.Towns__c !== areaInfo.town)
            return;
        if (!chomeInfo) {
            chomeInfo = new PriorityAreaInfo();
            chomeInfo.prefectures = record.Prefectures__c;
            chomeInfo.city = record.City__c;
            chomeInfo.townsAddress = record.WeeklyReportAnalysisChome__c;
            areaInfo.chomeList.set(chomeInfo.townsAddress, chomeInfo);
        }
        if (setNum === 1) {
            chomeInfo.countNum1++;
        }
        else {
            chomeInfo.countNum2++;
        }
    };
    AnalysReportService.prototype.outPutPriorityArea = function (workSheet, formatSheet) {
        var e_9, _a, e_10, _b;
        var optAreaMap = this.optAreaMap;
        if (!optAreaMap)
            return;
        var titleFormat = formatSheet.getCell("A6");
        var cityFormat = [];
        formatSheet.getRow(7).eachCell(function (cell) {
            cityFormat.push(cell);
        });
        var totalFormat = [];
        formatSheet.getRow(9).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        this.setPaTitle(workSheet, formatSheet, 13);
        var title = "";
        try {
            for (var _c = tslib_1.__values(optAreaMap.values()), _d = _c.next(); !_d.done; _d = _c.next()) {
                var areaInfo = _d.value;
                if (title != areaInfo.city) {
                    var targetCell = workSheet.getCell(this.PageInfo.rowIndex, this.PageInfo.colIndex);
                    targetCell.style = titleFormat.style;
                    targetCell.value = areaInfo.city;
                    this.PageInfo.runIndex();
                    title = areaInfo.city;
                }
                if (areaInfo.chomeList) {
                    try {
                        for (var _e = (e_10 = void 0, tslib_1.__values(areaInfo.chomeList.values())), _f = _e.next(); !_f.done; _f = _e.next()) {
                            var chomeInfo = _f.value;
                            var cells = this.getCellsByFormat(workSheet, cityFormat);
                            cells[0].value = chomeInfo.townsAddress;
                            cells[1].merge(cells[0]);
                            cells[2].value = chomeInfo.countNum1;
                            cells[3].value = areaInfo.countNum1 === 0 ? 0 : chomeInfo.countNum1 / areaInfo.countNum1;
                            cells[4].value = chomeInfo.countNum2;
                            cells[5].value = areaInfo.countNum2 === 0 ? 0 : chomeInfo.countNum2 / areaInfo.countNum2;
                            this.PageInfo.runIndex();
                            cells[6].value = chomeInfo.total;
                            cells[7].value = areaInfo.total === 0 ? 0 : chomeInfo.total / areaInfo.total;
                        }
                    }
                    catch (e_10_1) { e_10 = { error: e_10_1 }; }
                    finally {
                        try {
                            if (_f && !_f.done && (_b = _e.return)) _b.call(_e);
                        }
                        finally { if (e_10) throw e_10.error; }
                    }
                    this.PageInfo.runIndex();
                    this.optTotalForArea(areaInfo, workSheet, totalFormat);
                }
                else {
                    this.optTotalForArea(areaInfo, workSheet, totalFormat);
                }
            }
        }
        catch (e_9_1) { e_9 = { error: e_9_1 }; }
        finally {
            try {
                if (_d && !_d.done && (_a = _c.return)) _a.call(_c);
            }
            finally { if (e_9) throw e_9.error; }
        }
    };
    AnalysReportService.prototype.sumAge = function (record, optAgeMap) {
        if (!optAgeMap || !record.ResponsingAge__c)
            return;
        AgeInfo.ageTotalInfo.total++;
        var agePeriod = AgeInfo.getAgePeriod(record.ResponsingAge__c);
        var ageInfo = optAgeMap.get(agePeriod);
        if (!ageInfo) {
            ageInfo = new AgeInfo();
            ageInfo.agePeriod = agePeriod;
            optAgeMap.set(agePeriod, ageInfo);
        }
        ageInfo.total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            ageInfo.countNum1++;
            AgeInfo.ageTotalInfo.sum1Total++;
        }
        else {
            ageInfo.countNum2++;
            AgeInfo.ageTotalInfo.sum2Total++;
        }
    };
    AnalysReportService.prototype.outPutAgeInfo = function (workSheet, formatSheet) {
        var e_11, _a;
        if (!this.optAgeMap)
            return;
        var infoFormat = [];
        formatSheet.getRow(7).eachCell(function (cell) {
            infoFormat.push(cell);
        });
        this.setPaTitle(workSheet, formatSheet, 18);
        try {
            for (var _b = tslib_1.__values(this.optAgeMap.values()), _c = _b.next(); !_c.done; _c = _b.next()) {
                var ageInfo = _c.value;
                var cells_3 = this.getCellsByFormat(workSheet, infoFormat);
                cells_3[0].value = ageInfo.agePeriod;
                cells_3[1].merge(cells_3[0]);
                cells_3[2].value = ageInfo.countNum1;
                cells_3[3].value = ageInfo.countNum1 === 0 ? 0 : ageInfo.countNum1 / ageInfo.countNum1;
                cells_3[4].value = ageInfo.countNum2;
                cells_3[5].value = ageInfo.countNum2 === 0 ? 0 : ageInfo.countNum2 / ageInfo.countNum2;
                this.PageInfo.runIndex();
                cells_3[6].value = ageInfo.total;
                cells_3[7].value = ageInfo.total === 0 ? 0 : ageInfo.total / AgeInfo.ageTotalInfo.total;
            }
        }
        catch (e_11_1) { e_11 = { error: e_11_1 }; }
        finally {
            try {
                if (_c && !_c.done && (_a = _b.return)) _a.call(_b);
            }
            finally { if (e_11) throw e_11.error; }
        }
        this.PageInfo.runIndex();
        var totalFormat = [];
        formatSheet.getRow(9).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = AgeInfo.ageTotalInfo.sum1Total;
        cells[3].merge(cells[2]);
        cells[4].value = AgeInfo.ageTotalInfo.sum2Total;
        cells[5].merge(cells[4]);
        cells[6].value = AgeInfo.ageTotalInfo.total;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.sumFamilies = function (record, optFamiliesMap) {
        if (!optFamiliesMap || !record.NumberPeoplePlanMovein__c)
            return;
        FamiliesInfo.TotalInfo.total++;
        var faInfo = optFamiliesMap.get(record.NumberPeoplePlanMovein__c);
        if (!faInfo) {
            faInfo = new FamiliesInfo();
            faInfo.peopleNumber = record.NumberPeoplePlanMovein__c;
            optFamiliesMap.set(record.NumberPeoplePlanMovein__c, faInfo);
        }
        faInfo.total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            faInfo.countNum1++;
            FamiliesInfo.TotalInfo.sum1Total++;
        }
        else {
            faInfo.countNum2++;
            FamiliesInfo.TotalInfo.sum2Total++;
        }
    };
    AnalysReportService.prototype.outPutFamilieInfo = function (workSheet, formatSheet) {
        var e_12, _a;
        if (!this.optFamiliesMap)
            return;
        var infoFormat = [];
        formatSheet.getRow(7).eachCell(function (cell) {
            infoFormat.push(cell);
        });
        this.setPaTitle(workSheet, formatSheet, 23);
        try {
            for (var _b = tslib_1.__values(this.optFamiliesMap.values()), _c = _b.next(); !_c.done; _c = _b.next()) {
                var faInfo = _c.value;
                var cells_4 = this.getCellsByFormat(workSheet, infoFormat);
                cells_4[0].value = faInfo.peopleNumber + "人";
                cells_4[1].merge(cells_4[0]);
                cells_4[2].value = faInfo.countNum1;
                cells_4[3].value = faInfo.countNum1 === 0 ? 0 : faInfo.countNum1 / faInfo.countNum1;
                cells_4[4].value = faInfo.countNum2;
                cells_4[5].value = faInfo.countNum2 === 0 ? 0 : faInfo.countNum2 / faInfo.countNum2;
                this.PageInfo.runIndex();
                cells_4[6].value = faInfo.total;
                cells_4[7].value = faInfo.total === 0 ? 0 : faInfo.total / FamiliesInfo.TotalInfo.total;
            }
        }
        catch (e_12_1) { e_12 = { error: e_12_1 }; }
        finally {
            try {
                if (_c && !_c.done && (_a = _b.return)) _a.call(_b);
            }
            finally { if (e_12) throw e_12.error; }
        }
        this.PageInfo.runIndex();
        var totalFormat = [];
        formatSheet.getRow(9).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = FamiliesInfo.TotalInfo.sum1Total;
        cells[3].merge(cells[2]);
        cells[4].value = FamiliesInfo.TotalInfo.sum2Total;
        cells[5].merge(cells[4]);
        cells[6].value = FamiliesInfo.TotalInfo.total;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.sumLivingStyle = function (record, optLivingStyleMap) {
        if (!optLivingStyleMap || !record.LivingStyle__c)
            return;
        LivingStyleInfo.TotalInfo.total++;
        var faInfo = optLivingStyleMap.get(record.LivingStyle__c);
        if (!faInfo) {
            faInfo = new LivingStyleInfo();
            faInfo.name = record.LivingStyle__c;
            optLivingStyleMap.set(record.LivingStyle__c, faInfo);
        }
        faInfo.total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            faInfo.countNum1++;
            LivingStyleInfo.TotalInfo.sum1Total++;
        }
        else {
            faInfo.countNum2++;
            LivingStyleInfo.TotalInfo.sum2Total++;
        }
    };
    AnalysReportService.prototype.outPutLivingStyle = function (workSheet, formatSheet) {
        var e_13, _a;
        if (!this.optLivingStyleMap)
            return;
        var infoFormat = [];
        formatSheet.getRow(7).eachCell(function (cell) {
            infoFormat.push(cell);
        });
        this.setPaTitle(workSheet, formatSheet, 58);
        try {
            for (var _b = tslib_1.__values(this.optLivingStyleMap.values()), _c = _b.next(); !_c.done; _c = _b.next()) {
                var faInfo = _c.value;
                var cells_5 = this.getCellsByFormat(workSheet, infoFormat);
                cells_5[0].value = faInfo.name;
                cells_5[1].merge(cells_5[0]);
                cells_5[2].value = faInfo.countNum1;
                cells_5[3].value = faInfo.countNum1 === 0 ? 0 : faInfo.countNum1 / faInfo.countNum1;
                cells_5[4].value = faInfo.countNum2;
                cells_5[5].value = faInfo.countNum2 === 0 ? 0 : faInfo.countNum2 / faInfo.countNum2;
                this.PageInfo.runIndex();
                cells_5[6].value = faInfo.total;
                cells_5[7].value = faInfo.total === 0 ? 0 : faInfo.total / LivingStyleInfo.TotalInfo.total;
            }
        }
        catch (e_13_1) { e_13 = { error: e_13_1 }; }
        finally {
            try {
                if (_c && !_c.done && (_a = _b.return)) _a.call(_b);
            }
            finally { if (e_13) throw e_13.error; }
        }
        this.PageInfo.runIndex();
        var totalFormat = [];
        formatSheet.getRow(9).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = LivingStyleInfo.TotalInfo.sum1Total;
        cells[3].merge(cells[2]);
        cells[4].value = LivingStyleInfo.TotalInfo.sum2Total;
        cells[5].merge(cells[4]);
        cells[6].value = LivingStyleInfo.TotalInfo.total;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.sumReplacement = function (record, optReplaceMap) {
        if (!optReplaceMap || !record.ReplacementSchedule__c)
            return;
        ReplacementScheduleInfo.TotalInfo.total++;
        var faInfo = optReplaceMap.get(record.ReplacementSchedule__c);
        if (!faInfo) {
            faInfo = new ReplacementScheduleInfo();
            faInfo.name = record.ReplacementSchedule__c;
            optReplaceMap.set(record.ReplacementSchedule__c, faInfo);
        }
        faInfo.total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            faInfo.countNum1++;
            ReplacementScheduleInfo.TotalInfo.sum1Total++;
        }
        else {
            faInfo.countNum2++;
            ReplacementScheduleInfo.TotalInfo.sum2Total++;
        }
    };
    AnalysReportService.prototype.outPutReplacement = function (workSheet, formatSheet) {
        var e_14, _a;
        if (!this.optReplaceMap)
            return;
        var infoFormat = [];
        formatSheet.getRow(7).eachCell(function (cell) {
            infoFormat.push(cell);
        });
        this.setPaTitle(workSheet, formatSheet, 63);
        try {
            for (var _b = tslib_1.__values(this.optReplaceMap.values()), _c = _b.next(); !_c.done; _c = _b.next()) {
                var faInfo = _c.value;
                var cells_6 = this.getCellsByFormat(workSheet, infoFormat);
                cells_6[0].value = faInfo.name;
                cells_6[1].merge(cells_6[0]);
                cells_6[2].value = faInfo.countNum1;
                cells_6[3].value = faInfo.countNum1 === 0 ? 0 : faInfo.countNum1 / faInfo.countNum1;
                cells_6[4].value = faInfo.countNum2;
                cells_6[5].value = faInfo.countNum2 === 0 ? 0 : faInfo.countNum2 / faInfo.countNum2;
                this.PageInfo.runIndex();
                cells_6[6].value = faInfo.total;
                cells_6[7].value = faInfo.total === 0 ? 0 : faInfo.total / ReplacementScheduleInfo.TotalInfo.total;
            }
        }
        catch (e_14_1) { e_14 = { error: e_14_1 }; }
        finally {
            try {
                if (_c && !_c.done && (_a = _b.return)) _a.call(_b);
            }
            finally { if (e_14) throw e_14.error; }
        }
        this.PageInfo.runIndex();
        var totalFormat = [];
        formatSheet.getRow(9).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = ReplacementScheduleInfo.TotalInfo.sum1Total;
        cells[3].merge(cells[2]);
        cells[4].value = ReplacementScheduleInfo.TotalInfo.sum2Total;
        cells[5].merge(cells[4]);
        cells[6].value = ReplacementScheduleInfo.TotalInfo.total;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.sumReplacementProperty = function (record, optReplacePropertyMap) {
        if (!optReplacePropertyMap || !record.ReplacementProperty__c)
            return;
        ReplacementPropertyInfo.TotalInfo.total++;
        var faInfo = optReplacePropertyMap.get(record.ReplacementProperty__c);
        if (!faInfo) {
            faInfo = new ReplacementPropertyInfo();
            faInfo.name = record.ReplacementProperty__c;
            optReplacePropertyMap.set(record.ReplacementProperty__c, faInfo);
        }
        faInfo.total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            faInfo.countNum1++;
            ReplacementPropertyInfo.TotalInfo.sum1Total++;
        }
        else {
            faInfo.countNum2++;
            ReplacementPropertyInfo.TotalInfo.sum2Total++;
        }
    };
    AnalysReportService.prototype.outPutReplacementProperty = function (workSheet, formatSheet) {
        var e_15, _a;
        if (!this.optReplacePropertyMap)
            return;
        var infoFormat = [];
        formatSheet.getRow(7).eachCell(function (cell) {
            infoFormat.push(cell);
        });
        this.setPaTitle(workSheet, formatSheet, 68);
        try {
            for (var _b = tslib_1.__values(this.optReplacePropertyMap.values()), _c = _b.next(); !_c.done; _c = _b.next()) {
                var faInfo = _c.value;
                var cells_7 = this.getCellsByFormat(workSheet, infoFormat);
                cells_7[0].value = faInfo.name;
                cells_7[1].merge(cells_7[0]);
                cells_7[2].value = faInfo.countNum1;
                cells_7[3].value = faInfo.countNum1 === 0 ? 0 : faInfo.countNum1 / faInfo.countNum1;
                cells_7[4].value = faInfo.countNum2;
                cells_7[5].value = faInfo.countNum2 === 0 ? 0 : faInfo.countNum2 / faInfo.countNum2;
                this.PageInfo.runIndex();
                cells_7[6].value = faInfo.total;
                cells_7[7].value = faInfo.total === 0 ? 0 : faInfo.total / ReplacementPropertyInfo.TotalInfo.total;
            }
        }
        catch (e_15_1) { e_15 = { error: e_15_1 }; }
        finally {
            try {
                if (_c && !_c.done && (_a = _b.return)) _a.call(_b);
            }
            finally { if (e_15) throw e_15.error; }
        }
        this.PageInfo.runIndex();
        var totalFormat = [];
        formatSheet.getRow(9).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = ReplacementPropertyInfo.TotalInfo.sum1Total;
        cells[3].merge(cells[2]);
        cells[4].value = ReplacementPropertyInfo.TotalInfo.sum2Total;
        cells[5].merge(cells[4]);
        cells[6].value = ReplacementPropertyInfo.TotalInfo.total;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.sumProfession = function (record, optProfessionMap) {
        if (!optProfessionMap || !record.Profession__c)
            return;
        ProfessionInfo.TotalInfo.total++;
        var faInfo = optProfessionMap.get(record.Profession__c);
        if (!faInfo) {
            faInfo = new ProfessionInfo();
            faInfo.name = record.Profession__c;
            optProfessionMap.set(record.Profession__c, faInfo);
        }
        faInfo.total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            faInfo.countNum1++;
            ProfessionInfo.TotalInfo.sum1Total++;
        }
        else {
            faInfo.countNum2++;
            ProfessionInfo.TotalInfo.sum2Total++;
        }
    };
    AnalysReportService.prototype.outPutProfession = function (workSheet, formatSheet) {
        var e_16, _a;
        if (!this.optProfessionMap)
            return;
        var infoFormat = [];
        formatSheet.getRow(7).eachCell(function (cell) {
            infoFormat.push(cell);
        });
        this.setPaTitle(workSheet, formatSheet, 73);
        try {
            for (var _b = tslib_1.__values(this.optProfessionMap.values()), _c = _b.next(); !_c.done; _c = _b.next()) {
                var faInfo = _c.value;
                var cells_8 = this.getCellsByFormat(workSheet, infoFormat);
                cells_8[0].value = faInfo.name;
                cells_8[1].merge(cells_8[0]);
                cells_8[2].value = faInfo.countNum1;
                cells_8[3].value = faInfo.countNum1 === 0 ? 0 : faInfo.countNum1 / faInfo.countNum1;
                cells_8[4].value = faInfo.countNum2;
                cells_8[5].value = faInfo.countNum2 === 0 ? 0 : faInfo.countNum2 / faInfo.countNum2;
                this.PageInfo.runIndex();
                cells_8[6].value = faInfo.total;
                cells_8[7].value = faInfo.total === 0 ? 0 : faInfo.total / ProfessionInfo.TotalInfo.total;
            }
        }
        catch (e_16_1) { e_16 = { error: e_16_1 }; }
        finally {
            try {
                if (_c && !_c.done && (_a = _b.return)) _a.call(_b);
            }
            finally { if (e_16) throw e_16.error; }
        }
        this.PageInfo.runIndex();
        var totalFormat = [];
        formatSheet.getRow(9).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = ProfessionInfo.TotalInfo.sum1Total;
        cells[3].merge(cells[2]);
        cells[4].value = ProfessionInfo.TotalInfo.sum2Total;
        cells[5].merge(cells[4]);
        cells[6].value = ProfessionInfo.TotalInfo.total;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.sumParkingRequest = function (record, optParkingRequestMap) {
        if (!optParkingRequestMap || !record.ParkingRequest__c)
            return;
        ParkingRequestInfo.TotalInfo.total++;
        var faInfo = optParkingRequestMap.get(record.ParkingRequest__c);
        if (!faInfo) {
            faInfo = new ParkingRequestInfo();
            faInfo.name = record.ParkingRequest__c;
            optParkingRequestMap.set(record.ParkingRequest__c, faInfo);
        }
        faInfo.total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            faInfo.countNum1++;
            ParkingRequestInfo.TotalInfo.sum1Total++;
        }
        else {
            faInfo.countNum2++;
            ParkingRequestInfo.TotalInfo.sum2Total++;
        }
    };
    AnalysReportService.prototype.outPutParkingRequest = function (workSheet, formatSheet) {
        var e_17, _a;
        if (!this.optParkingRequestMap)
            return;
        var infoFormat = [];
        formatSheet.getRow(7).eachCell(function (cell) {
            infoFormat.push(cell);
        });
        this.setPaTitle(workSheet, formatSheet, 83);
        try {
            for (var _b = tslib_1.__values(this.optParkingRequestMap.values()), _c = _b.next(); !_c.done; _c = _b.next()) {
                var faInfo = _c.value;
                var cells_9 = this.getCellsByFormat(workSheet, infoFormat);
                cells_9[0].value = faInfo.name;
                cells_9[1].merge(cells_9[0]);
                cells_9[2].value = faInfo.countNum1;
                cells_9[3].value = faInfo.countNum1 === 0 ? 0 : faInfo.countNum1 / faInfo.countNum1;
                cells_9[4].value = faInfo.countNum2;
                cells_9[5].value = faInfo.countNum2 === 0 ? 0 : faInfo.countNum2 / faInfo.countNum2;
                this.PageInfo.runIndex();
                cells_9[6].value = faInfo.total;
                cells_9[7].value = faInfo.total === 0 ? 0 : faInfo.total / ParkingRequestInfo.TotalInfo.total;
            }
        }
        catch (e_17_1) { e_17 = { error: e_17_1 }; }
        finally {
            try {
                if (_c && !_c.done && (_a = _b.return)) _a.call(_b);
            }
            finally { if (e_17) throw e_17.error; }
        }
        this.PageInfo.runIndex();
        var totalFormat = [];
        formatSheet.getRow(9).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = ParkingRequestInfo.TotalInfo.sum1Total;
        cells[3].merge(cells[2]);
        cells[4].value = ParkingRequestInfo.TotalInfo.sum2Total;
        cells[5].merge(cells[4]);
        cells[6].value = ParkingRequestInfo.TotalInfo.total;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.sumParkingNumber = function (record, optParkingNumberMap) {
        if (!optParkingNumberMap || !record.ParkingNumber__c)
            return;
        ParkingNumberInfo.TotalInfo.total++;
        var faInfo = optParkingNumberMap.get(record.ParkingNumber__c + "台");
        if (!faInfo) {
            faInfo = new ParkingNumberInfo();
            faInfo.name = record.ParkingNumber__c + "台";
            optParkingNumberMap.set(faInfo.name, faInfo);
        }
        faInfo.total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            faInfo.countNum1++;
            ParkingNumberInfo.TotalInfo.sum1Total++;
        }
        else {
            faInfo.countNum2++;
            ParkingNumberInfo.TotalInfo.sum2Total++;
        }
    };
    AnalysReportService.prototype.outPutParkingNumber = function (workSheet, formatSheet) {
        var e_18, _a;
        if (!this.optParkingNumberMap)
            return;
        var infoFormat = [];
        formatSheet.getRow(7).eachCell(function (cell) {
            infoFormat.push(cell);
        });
        this.setPaTitle(workSheet, formatSheet, 88);
        try {
            for (var _b = tslib_1.__values(this.optParkingNumberMap.values()), _c = _b.next(); !_c.done; _c = _b.next()) {
                var faInfo = _c.value;
                var cells_10 = this.getCellsByFormat(workSheet, infoFormat);
                cells_10[0].value = faInfo.name;
                cells_10[1].merge(cells_10[0]);
                cells_10[2].value = faInfo.countNum1;
                cells_10[3].value = faInfo.countNum1 === 0 ? 0 : faInfo.countNum1 / faInfo.countNum1;
                cells_10[4].value = faInfo.countNum2;
                cells_10[5].value = faInfo.countNum2 === 0 ? 0 : faInfo.countNum2 / faInfo.countNum2;
                this.PageInfo.runIndex();
                cells_10[6].value = faInfo.total;
                cells_10[7].value = faInfo.total === 0 ? 0 : faInfo.total / ParkingNumberInfo.TotalInfo.total;
            }
        }
        catch (e_18_1) { e_18 = { error: e_18_1 }; }
        finally {
            try {
                if (_c && !_c.done && (_a = _b.return)) _a.call(_b);
            }
            finally { if (e_18) throw e_18.error; }
        }
        this.PageInfo.runIndex();
        var totalFormat = [];
        formatSheet.getRow(9).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = ParkingNumberInfo.TotalInfo.sum1Total;
        cells[3].merge(cells[2]);
        cells[4].value = ParkingNumberInfo.TotalInfo.sum2Total;
        cells[5].merge(cells[4]);
        cells[6].value = ParkingNumberInfo.TotalInfo.total;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.setPaTitle = function (workSheet, formatSheet, formatCellRow, customTitle) {
        var _this = this;
        if (this.PageInfo.hasRow < 4) {
            this.PageInfo.runIndexByNum(this.PageInfo.hasRow);
        }
        var paTitle = formatSheet.getCell("A" + formatCellRow);
        workSheet.mergeCells(this.PageInfo.rowIndex, this.PageInfo.colIndex, this.PageInfo.rowIndex + 1, this.PageInfo.colIndex + 7);
        var paTitleCell = workSheet.getCell(this.PageInfo.rowIndex, this.PageInfo.colIndex);
        paTitleCell.style = paTitle.style;
        paTitleCell.value = customTitle ? customTitle : paTitle.value;
        this.PageInfo.runIndexByNum(2);
        var index = 0;
        var beforceVal;
        formatSheet.getRow(formatCellRow + 2).eachCell(function (cell) {
            var tgtCell = workSheet.getCell(_this.PageInfo.rowIndex, _this.PageInfo.colIndex + index);
            tgtCell.style = cell.style;
            tgtCell.value = cell.value;
            if (beforceVal && beforceVal === cell.value && index > 1) {
                tgtCell.merge(workSheet.getCell(_this.PageInfo.rowIndex, _this.PageInfo.colIndex + index - 1));
            }
            beforceVal = tgtCell.value;
            index++;
        });
        index = 0;
        beforceVal = undefined;
        this.PageInfo.runIndex();
        formatSheet.getRow(formatCellRow + 3).eachCell(function (cell) {
            var tgtCell = workSheet.getCell(_this.PageInfo.rowIndex, _this.PageInfo.colIndex + index);
            tgtCell.style = cell.style;
            tgtCell.value = cell.value;
            if (beforceVal && beforceVal === cell.value && index > 1) {
                tgtCell.merge(workSheet.getCell(_this.PageInfo.rowIndex, _this.PageInfo.colIndex + index - 1));
            }
            beforceVal = tgtCell.value;
            index++;
        });
        workSheet.mergeCells(this.PageInfo.rowIndex - 1, this.PageInfo.colIndex, this.PageInfo.rowIndex, this.PageInfo.colIndex + 1);
        workSheet.getCell(this.PageInfo.rowIndex, this.PageInfo.colIndex).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" }
        };
        this.PageInfo.runIndexByNum(2);
    };
    AnalysReportService.prototype.setMeTitle = function (workSheet, formatSheet, formatCellRow, customTitle, head) {
        var PageInfo = this.MediumPageInfo;
        var paTitle = formatSheet.getCell("A" + formatCellRow);
        workSheet.mergeCells(PageInfo.rowIndex, PageInfo.colIndex, PageInfo.rowIndex + 1, PageInfo.colIndex + 11);
        var paTitleCell = workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex);
        paTitleCell.style = paTitle.style;
        paTitleCell.value = customTitle ? customTitle : paTitle.value;
        PageInfo.runIndex(3);
        var index = 0;
        var beforceVal;
        formatSheet.getRow(formatCellRow + 3).eachCell(function (cell) {
            var tgtCell = workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex + index);
            tgtCell.style = cell.style;
            tgtCell.value = cell.value;
            if (beforceVal && beforceVal === cell.value && index > 1) {
                tgtCell.merge(workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex + index - 1));
            }
            beforceVal = tgtCell.value;
            index++;
        });
        if (head) {
            formatSheet.getCell(formatCellRow + 3, PageInfo.colIndex).value = head;
        }
        index = 0;
        beforceVal = undefined;
        PageInfo.runIndex();
        formatSheet.getRow(formatCellRow + 4).eachCell(function (cell) {
            var tgtCell = workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex + index);
            tgtCell.style = cell.style;
            tgtCell.value = cell.value;
            if (beforceVal && beforceVal === cell.value && index > 1) {
                tgtCell.merge(workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex + index - 1));
            }
            beforceVal = tgtCell.value;
            index++;
        });
        workSheet.mergeCells(PageInfo.rowIndex - 1, PageInfo.colIndex, PageInfo.rowIndex, PageInfo.colIndex + 1);
        workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" }
        };
        PageInfo.runIndex(2);
    };
    AnalysReportService.prototype.setOtherTitle = function (workSheet, formatSheet, formatCellRow, customTitle) {
        var PageInfo = this.OtherPageInfo;
        if (PageInfo.hasRow < 4) {
            PageInfo.runIndexByNum(PageInfo.hasRow + 1);
        }
        var paTitle = formatSheet.getCell("A" + formatCellRow);
        workSheet.mergeCells(PageInfo.rowIndex, PageInfo.colIndex, PageInfo.rowIndex + 1, PageInfo.colIndex + 7);
        var paTitleCell = workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex);
        paTitleCell.style = paTitle.style;
        paTitleCell.value = customTitle ? customTitle : paTitle.value;
        PageInfo.runIndexByNum(2);
        var index = 0;
        var beforceVal;
        formatSheet.getRow(formatCellRow + 2).eachCell(function (cell) {
            var tgtCell = workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex + index);
            tgtCell.style = cell.style;
            tgtCell.value = cell.value;
            if (beforceVal && beforceVal === cell.value && index > 1) {
                tgtCell.merge(workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex + index - 1));
            }
            beforceVal = tgtCell.value;
            index++;
        });
        index = 0;
        beforceVal = undefined;
        PageInfo.runIndex();
        formatSheet.getRow(formatCellRow + 3).eachCell(function (cell) {
            var tgtCell = workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex + index);
            tgtCell.style = cell.style;
            tgtCell.value = cell.value;
            if (beforceVal && beforceVal === cell.value && index > 1) {
                tgtCell.merge(workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex + index - 1));
            }
            beforceVal = tgtCell.value;
            index++;
        });
        workSheet.mergeCells(PageInfo.rowIndex - 1, PageInfo.colIndex, PageInfo.rowIndex, PageInfo.colIndex + 1);
        workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" }
        };
        PageInfo.runIndexByNum(2);
    };
    AnalysReportService.prototype.sumAnnualIncome = function (record, optAnnualIncomeMap) {
        if (!optAnnualIncomeMap || !record.GenerationMainAnnualIncome__c)
            return;
        AnnualIncomeInfo.TotalInfo.total++;
        var aggregateInfo = this.analysisRptConfig.aggregatePitch.annualIncome;
        var key;
        var amountStart = 0;
        if (record.GenerationMainAnnualIncome__c < aggregateInfo.start) {
            amountStart = aggregateInfo.start - 1;
            key = "～" + amountStart + "万円";
        }
        else if (record.GenerationMainAnnualIncome__c >= aggregateInfo.end) {
            amountStart = aggregateInfo.end;
            key = amountStart + "万円～";
        }
        else {
            amountStart = ~~(record.GenerationMainAnnualIncome__c / aggregateInfo.unit) * aggregateInfo.unit;
            key = amountStart + "万円～";
        }
        var alInfo = optAnnualIncomeMap.get(key);
        if (!alInfo) {
            alInfo = new AnnualIncomeInfo();
            alInfo.aggregateName = key;
            alInfo.amountStart = amountStart;
            optAnnualIncomeMap.set(key, alInfo);
        }
        alInfo.total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            alInfo.countNum1++;
            AnnualIncomeInfo.TotalInfo.sum1Total++;
        }
        else {
            alInfo.countNum2++;
            AnnualIncomeInfo.TotalInfo.sum2Total++;
        }
    };
    AnalysReportService.prototype.outPutAnnualIncome = function (workSheet, formatSheet) {
        var e_19, _a, e_20, _b;
        if (!this.optAnnualIncomeMap)
            return;
        var infoFormat = [];
        formatSheet.getRow(7).eachCell(function (cell) {
            infoFormat.push(cell);
        });
        this.setPaTitle(workSheet, formatSheet, 28, "\u4E16\u5E2F\u4E3B\u5E74\u53CE(" + this.analysisRptConfig.aggregatePitch.annualIncome.start + "\u4E07\u5186\uFF5E" + this.analysisRptConfig.aggregatePitch.annualIncome.end + "\u4E07\u5186  " + this.analysisRptConfig.aggregatePitch.annualIncome.unit + "\u4E07\u5186\u6BCE)\u306E\u96C6\u8A08");
        var sortList = [];
        try {
            for (var _c = tslib_1.__values(this.optAnnualIncomeMap.values()), _d = _c.next(); !_d.done; _d = _c.next()) {
                var aiInfo = _d.value;
                sortList.push(aiInfo);
            }
        }
        catch (e_19_1) { e_19 = { error: e_19_1 }; }
        finally {
            try {
                if (_d && !_d.done && (_a = _c.return)) _a.call(_c);
            }
            finally { if (e_19) throw e_19.error; }
        }
        sortList.sort(function (a, b) {
            return a.amountStart - b.amountStart;
        });
        try {
            for (var sortList_1 = tslib_1.__values(sortList), sortList_1_1 = sortList_1.next(); !sortList_1_1.done; sortList_1_1 = sortList_1.next()) {
                var aiInfo = sortList_1_1.value;
                var cells_11 = this.getCellsByFormat(workSheet, infoFormat);
                cells_11[0].value = aiInfo.aggregateName;
                cells_11[1].merge(cells_11[0]);
                cells_11[2].value = aiInfo.countNum1;
                cells_11[3].value = aiInfo.countNum1 === 0 ? 0 : aiInfo.countNum1 / aiInfo.countNum1;
                cells_11[4].value = aiInfo.countNum2;
                cells_11[5].value = aiInfo.countNum2 === 0 ? 0 : aiInfo.countNum2 / aiInfo.countNum2;
                cells_11[6].value = aiInfo.total;
                cells_11[7].value = aiInfo.total === 0 ? 0 : aiInfo.total / AnnualIncomeInfo.TotalInfo.total;
                this.PageInfo.runIndex();
            }
        }
        catch (e_20_1) { e_20 = { error: e_20_1 }; }
        finally {
            try {
                if (sortList_1_1 && !sortList_1_1.done && (_b = sortList_1.return)) _b.call(sortList_1);
            }
            finally { if (e_20) throw e_20.error; }
        }
        this.PageInfo.runIndex();
        var totalFormat = [];
        formatSheet.getRow(9).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = AnnualIncomeInfo.TotalInfo.sum1Total;
        cells[3].merge(cells[2]);
        cells[4].value = AnnualIncomeInfo.TotalInfo.sum2Total;
        cells[5].merge(cells[4]);
        cells[6].value = AnnualIncomeInfo.TotalInfo.total;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.sumOwnResources = function (record, optOwnResources) {
        if (!optOwnResources || !record.OwnResources__c)
            return;
        OwnResourcesInfo.TotalInfo.total++;
        var aggregateInfo = this.analysisRptConfig.aggregatePitch.ownResources;
        var key;
        var amountStart = 0;
        if (record.OwnResources__c < aggregateInfo.start) {
            amountStart = aggregateInfo.start - 1;
            key = "～" + amountStart + "万円";
        }
        else if (record.OwnResources__c >= aggregateInfo.end) {
            amountStart = aggregateInfo.end;
            key = amountStart + "万円～";
        }
        else {
            amountStart = ~~(record.OwnResources__c / aggregateInfo.unit) * aggregateInfo.unit;
            key = amountStart + "万円～";
        }
        var alInfo = optOwnResources.get(key);
        if (!alInfo) {
            alInfo = new OwnResourcesInfo();
            alInfo.aggregateName = key;
            alInfo.amountStart = amountStart;
            optOwnResources.set(key, alInfo);
        }
        alInfo.total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            alInfo.countNum1++;
            OwnResourcesInfo.TotalInfo.sum1Total++;
        }
        else {
            alInfo.countNum2++;
            OwnResourcesInfo.TotalInfo.sum2Total++;
        }
    };
    AnalysReportService.prototype.outPutOwnResources = function (workSheet, formatSheet) {
        var e_21, _a, e_22, _b;
        if (!this.optOwnResources)
            return;
        var infoFormat = [];
        formatSheet.getRow(7).eachCell(function (cell) {
            infoFormat.push(cell);
        });
        this.setPaTitle(workSheet, formatSheet, 43, "\u81EA\u5DF1\u8CC7\u91D1(" + this.analysisRptConfig.aggregatePitch.ownResources.start + "\u4E07\u5186\uFF5E" + this.analysisRptConfig.aggregatePitch.ownResources.end + "\u4E07\u5186  " + this.analysisRptConfig.aggregatePitch.ownResources.unit + "\u4E07\u5186\u6BCE)\u306E\u96C6\u8A08");
        var sortList = [];
        try {
            for (var _c = tslib_1.__values(this.optOwnResources.values()), _d = _c.next(); !_d.done; _d = _c.next()) {
                var aiInfo = _d.value;
                sortList.push(aiInfo);
            }
        }
        catch (e_21_1) { e_21 = { error: e_21_1 }; }
        finally {
            try {
                if (_d && !_d.done && (_a = _c.return)) _a.call(_c);
            }
            finally { if (e_21) throw e_21.error; }
        }
        sortList.sort(function (a, b) {
            return a.amountStart - b.amountStart;
        });
        try {
            for (var sortList_2 = tslib_1.__values(sortList), sortList_2_1 = sortList_2.next(); !sortList_2_1.done; sortList_2_1 = sortList_2.next()) {
                var aiInfo = sortList_2_1.value;
                var cells_12 = this.getCellsByFormat(workSheet, infoFormat);
                cells_12[0].value = aiInfo.aggregateName;
                cells_12[1].merge(cells_12[0]);
                cells_12[2].value = aiInfo.countNum1;
                cells_12[3].value = aiInfo.countNum1 === 0 ? 0 : aiInfo.countNum1 / aiInfo.countNum1;
                cells_12[4].value = aiInfo.countNum2;
                cells_12[5].value = aiInfo.countNum2 === 0 ? 0 : aiInfo.countNum2 / aiInfo.countNum2;
                cells_12[6].value = aiInfo.total;
                cells_12[7].value = aiInfo.total === 0 ? 0 : aiInfo.total / OwnResourcesInfo.TotalInfo.total;
                this.PageInfo.runIndex();
            }
        }
        catch (e_22_1) { e_22 = { error: e_22_1 }; }
        finally {
            try {
                if (sortList_2_1 && !sortList_2_1.done && (_b = sortList_2.return)) _b.call(sortList_2);
            }
            finally { if (e_22) throw e_22.error; }
        }
        this.PageInfo.runIndex();
        var totalFormat = [];
        formatSheet.getRow(9).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = OwnResourcesInfo.TotalInfo.sum1Total;
        cells[3].merge(cells[2]);
        cells[4].value = OwnResourcesInfo.TotalInfo.sum2Total;
        cells[5].merge(cells[4]);
        cells[6].value = OwnResourcesInfo.TotalInfo.total;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.sumBudget = function (record, optBudgetMap) {
        if (!optBudgetMap || !record.OwnResources__c)
            return;
        BudgetInfo.TotalInfo.total++;
        var aggregateInfo = this.analysisRptConfig.aggregatePitch.budget;
        var key;
        var amountStart = 0;
        if (record.OwnResources__c < aggregateInfo.start) {
            return;
        }
        else if (record.OwnResources__c >= aggregateInfo.end) {
            amountStart = aggregateInfo.end;
            key = amountStart + "万円～";
        }
        else {
            amountStart = ~~(record.OwnResources__c / aggregateInfo.unit) * aggregateInfo.unit;
            key = amountStart + "万円～";
        }
        var bgInfo = optBudgetMap.get(key);
        if (!bgInfo) {
            bgInfo = new BudgetInfo();
            bgInfo.aggregateName = key;
            bgInfo.amountStart = amountStart;
            optBudgetMap.set(key, bgInfo);
        }
        bgInfo.total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            bgInfo.countNum1++;
            BudgetInfo.TotalInfo.sum1Total++;
        }
        else {
            bgInfo.countNum2++;
            BudgetInfo.TotalInfo.sum2Total++;
        }
    };
    AnalysReportService.prototype.outPutBudget = function (workSheet, formatSheet) {
        var e_23, _a, e_24, _b;
        if (!this.optBudgetMap)
            return;
        var infoFormat = [];
        formatSheet.getRow(7).eachCell(function (cell) {
            infoFormat.push(cell);
        });
        this.setPaTitle(workSheet, formatSheet, 48, "\u4E88\u7B97(" + this.analysisRptConfig.aggregatePitch.budget.start + "\u4E07\u5186\uFF5E" + this.analysisRptConfig.aggregatePitch.budget.end + "\u4E07\u5186  " + this.analysisRptConfig.aggregatePitch.budget.unit + "\u4E07\u5186\u6BCE)\u306E\u96C6\u8A08");
        var sortList = [];
        try {
            for (var _c = tslib_1.__values(this.optBudgetMap.values()), _d = _c.next(); !_d.done; _d = _c.next()) {
                var aiInfo = _d.value;
                sortList.push(aiInfo);
            }
        }
        catch (e_23_1) { e_23 = { error: e_23_1 }; }
        finally {
            try {
                if (_d && !_d.done && (_a = _c.return)) _a.call(_c);
            }
            finally { if (e_23) throw e_23.error; }
        }
        sortList.sort(function (a, b) {
            return a.amountStart - b.amountStart;
        });
        try {
            for (var sortList_3 = tslib_1.__values(sortList), sortList_3_1 = sortList_3.next(); !sortList_3_1.done; sortList_3_1 = sortList_3.next()) {
                var aiInfo = sortList_3_1.value;
                var cells_13 = this.getCellsByFormat(workSheet, infoFormat);
                cells_13[0].value = aiInfo.aggregateName;
                cells_13[1].merge(cells_13[0]);
                cells_13[2].value = aiInfo.countNum1;
                cells_13[3].value = aiInfo.countNum1 === 0 ? 0 : aiInfo.countNum1 / aiInfo.countNum1;
                cells_13[4].value = aiInfo.countNum2;
                cells_13[5].value = aiInfo.countNum2 === 0 ? 0 : aiInfo.countNum2 / aiInfo.countNum2;
                cells_13[6].value = aiInfo.total;
                cells_13[7].value = aiInfo.total === 0 ? 0 : aiInfo.total / BudgetInfo.TotalInfo.total;
                this.PageInfo.runIndex();
            }
        }
        catch (e_24_1) { e_24 = { error: e_24_1 }; }
        finally {
            try {
                if (sortList_3_1 && !sortList_3_1.done && (_b = sortList_3.return)) _b.call(sortList_3);
            }
            finally { if (e_24) throw e_24.error; }
        }
        this.PageInfo.runIndex();
        var totalFormat = [];
        formatSheet.getRow(9).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = BudgetInfo.TotalInfo.sum1Total;
        cells[3].merge(cells[2]);
        cells[4].value = BudgetInfo.TotalInfo.sum2Total;
        cells[5].merge(cells[4]);
        cells[6].value = BudgetInfo.TotalInfo.total;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.sumDesiredArea = function (record, optDesiredAreaMap) {
        if (!optDesiredAreaMap || !record.DesiredArea__c)
            return;
        DesiredAreaInfo.TotalInfo.total++;
        var aggregateInfo = this.analysisRptConfig.aggregatePitch.desiredArea;
        var key;
        var amountStart = 0;
        if (record.DesiredArea__c < aggregateInfo.start) {
            return;
        }
        else if (record.DesiredArea__c >= aggregateInfo.end) {
            amountStart = aggregateInfo.end;
            key = amountStart + "㎡～";
        }
        else {
            amountStart = ~~(record.DesiredArea__c / aggregateInfo.unit) * aggregateInfo.unit;
            key = amountStart + "㎡～";
        }
        var optInfo = optDesiredAreaMap.get(key);
        if (!optInfo) {
            optInfo = new DesiredAreaInfo();
            optInfo.aggregateName = key;
            optInfo.numStart = amountStart;
            optDesiredAreaMap.set(key, optInfo);
        }
        optInfo.total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            optInfo.countNum1++;
            DesiredAreaInfo.TotalInfo.sum1Total++;
        }
        else {
            optInfo.countNum2++;
            DesiredAreaInfo.TotalInfo.sum2Total++;
        }
    };
    AnalysReportService.prototype.outputDesiredArea = function (workSheet, formatSheet) {
        var e_25, _a, e_26, _b;
        if (!this.optDesiredAreaMap)
            return;
        var infoFormat = [];
        formatSheet.getRow(7).eachCell(function (cell) {
            infoFormat.push(cell);
        });
        this.setPaTitle(workSheet, formatSheet, 53);
        var sortList = [];
        try {
            for (var _c = tslib_1.__values(this.optDesiredAreaMap.values()), _d = _c.next(); !_d.done; _d = _c.next()) {
                var aiInfo = _d.value;
                sortList.push(aiInfo);
            }
        }
        catch (e_25_1) { e_25 = { error: e_25_1 }; }
        finally {
            try {
                if (_d && !_d.done && (_a = _c.return)) _a.call(_c);
            }
            finally { if (e_25) throw e_25.error; }
        }
        sortList.sort(function (a, b) {
            return a.numStart - b.numStart;
        });
        try {
            for (var sortList_4 = tslib_1.__values(sortList), sortList_4_1 = sortList_4.next(); !sortList_4_1.done; sortList_4_1 = sortList_4.next()) {
                var aiInfo = sortList_4_1.value;
                var cells_14 = this.getCellsByFormat(workSheet, infoFormat);
                cells_14[0].value = aiInfo.aggregateName;
                cells_14[1].merge(cells_14[0]);
                cells_14[2].value = aiInfo.countNum1;
                cells_14[3].value = aiInfo.countNum1 === 0 ? 0 : aiInfo.countNum1 / aiInfo.countNum1;
                cells_14[4].value = aiInfo.countNum2;
                cells_14[5].value = aiInfo.countNum2 === 0 ? 0 : aiInfo.countNum2 / aiInfo.countNum2;
                cells_14[6].value = aiInfo.total;
                cells_14[7].value = aiInfo.total === 0 ? 0 : aiInfo.total / DesiredAreaInfo.TotalInfo.total;
                this.PageInfo.runIndex();
            }
        }
        catch (e_26_1) { e_26 = { error: e_26_1 }; }
        finally {
            try {
                if (sortList_4_1 && !sortList_4_1.done && (_b = sortList_4.return)) _b.call(sortList_4);
            }
            finally { if (e_26) throw e_26.error; }
        }
        this.PageInfo.runIndex();
        var totalFormat = [];
        formatSheet.getRow(9).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = DesiredAreaInfo.TotalInfo.sum1Total;
        cells[3].merge(cells[2]);
        cells[4].value = DesiredAreaInfo.TotalInfo.sum2Total;
        cells[5].merge(cells[4]);
        cells[6].value = DesiredAreaInfo.TotalInfo.total;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.sumDesiredFloorPlan = function (record, plan, desiredFloorPlanInfo) {
        if (!desiredFloorPlanInfo)
            return;
        var DesiredFloorPlanVal;
        if (plan === "plan1") {
            DesiredFloorPlanVal = record.DesiredFloorPlan1__c;
        }
        else {
            DesiredFloorPlanVal = record.DesiredFloorPlan2__c;
        }
        if (!DesiredFloorPlanVal)
            return;
        desiredFloorPlanInfo[plan + "TotalInfo"].total++;
        var dfInfo = desiredFloorPlanInfo[plan].get(DesiredFloorPlanVal);
        if (!dfInfo) {
            dfInfo = new DesiredFloorPlan();
            dfInfo.name = DesiredFloorPlanVal;
            desiredFloorPlanInfo[plan].set(DesiredFloorPlanVal, dfInfo);
        }
        dfInfo.total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            dfInfo.countNum1++;
            desiredFloorPlanInfo[plan + "TotalInfo"].sum1Total++;
        }
        else {
            dfInfo.countNum2++;
            desiredFloorPlanInfo[plan + "TotalInfo"].sum2Total++;
        }
    };
    AnalysReportService.prototype.outDesiredFloorPlan = function (workSheet, formatSheet, plan) {
        var e_27, _a, e_28, _b;
        if (!this.desiredFloorPlanInfo)
            return;
        var infoFormat = [];
        formatSheet.getRow(7).eachCell(function (cell) {
            infoFormat.push(cell);
        });
        if (plan === "plan1") {
            this.setPaTitle(workSheet, formatSheet, 33);
        }
        else {
            this.setPaTitle(workSheet, formatSheet, 38);
        }
        var optMap = this.desiredFloorPlanInfo[plan];
        var sortList = [];
        try {
            for (var _c = tslib_1.__values(optMap.values()), _d = _c.next(); !_d.done; _d = _c.next()) {
                var aiInfo = _d.value;
                sortList.push(aiInfo);
            }
        }
        catch (e_27_1) { e_27 = { error: e_27_1 }; }
        finally {
            try {
                if (_d && !_d.done && (_a = _c.return)) _a.call(_c);
            }
            finally { if (e_27) throw e_27.error; }
        }
        sortList.sort(function (a, b) {
            return a.name.localeCompare(b.name);
        });
        try {
            for (var sortList_5 = tslib_1.__values(sortList), sortList_5_1 = sortList_5.next(); !sortList_5_1.done; sortList_5_1 = sortList_5.next()) {
                var faInfo = sortList_5_1.value;
                var cells_15 = this.getCellsByFormat(workSheet, infoFormat);
                cells_15[0].value = faInfo.name;
                cells_15[1].merge(cells_15[0]);
                cells_15[2].value = faInfo.countNum1;
                cells_15[3].value = faInfo.countNum1 === 0 ? 0 : faInfo.countNum1 / faInfo.countNum1;
                cells_15[4].value = faInfo.countNum2;
                cells_15[5].value = faInfo.countNum2 === 0 ? 0 : faInfo.countNum2 / faInfo.countNum2;
                this.PageInfo.runIndex();
                cells_15[6].value = faInfo.total;
                cells_15[7].value = faInfo.total === 0 ? 0 : faInfo.total / this.desiredFloorPlanInfo[plan + "TotalInfo"].total;
            }
        }
        catch (e_28_1) { e_28 = { error: e_28_1 }; }
        finally {
            try {
                if (sortList_5_1 && !sortList_5_1.done && (_b = sortList_5.return)) _b.call(sortList_5);
            }
            finally { if (e_28) throw e_28.error; }
        }
        this.PageInfo.runIndex();
        var totalFormat = [];
        formatSheet.getRow(9).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        var cells = this.getCellsByFormat(workSheet, totalFormat);
        cells[0].value = "\u5408\u8A08";
        cells[1].merge(cells[0]);
        cells[2].value = this.desiredFloorPlanInfo[plan + "TotalInfo"].sum1Total;
        cells[3].merge(cells[2]);
        cells[4].value = this.desiredFloorPlanInfo[plan + "TotalInfo"].sum2Total;
        cells[5].merge(cells[4]);
        cells[6].value = this.desiredFloorPlanInfo[plan + "TotalInfo"].total;
        cells[7].merge(cells[6]);
        this.PageInfo.runIndex();
        this.PageInfo.runIndex();
    };
    AnalysReportService.prototype.sumMedium1 = function (record) {
        if (!this.optMediumMap1)
            return;
        if (!record.MajorItems__c || !record.SubItem__c)
            return;
        var preInfo = this.optMediumMap1.get(record.MajorItems__c);
        if (!preInfo) {
            preInfo = new MediumInfo();
            preInfo.name = record.MajorItems__c;
            if (record.MajorItems__c === "インターネット") {
                preInfo.itemList = new Map();
            }
            this.optMediumMap1.set(preInfo.name, preInfo);
        }
        preInfo.total++;
        this.md1TotalInfo.Total++;
        var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            preInfo.countNum1++;
            this.md1TotalInfo.sum1Total++;
        }
        else {
            preInfo.countNum2++;
            this.md1TotalInfo.sum2Total++;
        }
        var itemName = record.MediumInternet__c ? record.MediumInternet__c : record.SubItem__c;
        if (preInfo.itemList && itemName) {
            var itemInfo = preInfo.itemList.get(itemName);
            if (!itemInfo) {
                itemInfo = new MediumInfo();
                itemInfo.name = itemName;
                preInfo.itemList.set(itemInfo.name, itemInfo);
            }
            if (setNum === 1) {
                itemInfo.countNum1++;
            }
            else {
                itemInfo.countNum2++;
            }
            itemInfo.total++;
        }
    };
    AnalysReportService.prototype.outPutMedium1 = function (workSheet, formatSheet) {
        var e_29, _a, e_30, _b;
        var optMap = this.optMediumMap1;
        if (!optMap)
            return;
        var PageInfo = this.MediumPageInfo;
        PageInfo.colNum = 0;
        var titleFormat = formatSheet.getCell("A100");
        var itemFormat = [];
        formatSheet.getRow(101).eachCell(function (cell) {
            itemFormat.push(cell);
        });
        var itemTotalFormat = [];
        formatSheet.getRow(105).eachCell({ includeEmpty: true }, function (cell) {
            itemTotalFormat.push(cell);
        });
        var totalFormat = [];
        formatSheet.getRow(103).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        try {
            for (var _c = tslib_1.__values(optMap.values()), _d = _c.next(); !_d.done; _d = _c.next()) {
                var preInfo = _d.value;
                if (preInfo.name !== "インターネット以外") {
                    var targetCell = workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex);
                    targetCell.style = titleFormat.style;
                    targetCell.value = preInfo.name;
                    PageInfo.runIndex();
                }
                if (preInfo.itemList && preInfo.itemList.size > 0) {
                    try {
                        for (var _e = (e_30 = void 0, tslib_1.__values(preInfo.itemList.values())), _f = _e.next(); !_f.done; _f = _e.next()) {
                            var itemInfo = _f.value;
                            var cells_16 = this.getCellsByFormat(workSheet, itemFormat, PageInfo);
                            cells_16[0].value = itemInfo.name;
                            cells_16[1].merge(cells_16[0]);
                            cells_16[2].merge(cells_16[0]);
                            cells_16[3].value = itemInfo.countNum1;
                            cells_16[4].value = preInfo.countNum1 === 0 ? 0 : itemInfo.countNum1 / preInfo.countNum1;
                            cells_16[5].value = preInfo.countNum1 === 0 ? 0 : itemInfo.countNum1 / preInfo.countNum1;
                            cells_16[6].value = itemInfo.countNum2;
                            cells_16[7].value = preInfo.countNum2 === 0 ? 0 : itemInfo.countNum2 / preInfo.countNum2;
                            cells_16[8].value = preInfo.countNum2 === 0 ? 0 : itemInfo.countNum2 / preInfo.countNum2;
                            cells_16[9].value = itemInfo.total;
                            cells_16[10].value = preInfo.total === 0 ? 0 : itemInfo.total / preInfo.total;
                            cells_16[11].value = preInfo.total === 0 ? 0 : itemInfo.total / preInfo.total;
                            PageInfo.runIndex();
                        }
                    }
                    catch (e_30_1) { e_30 = { error: e_30_1 }; }
                    finally {
                        try {
                            if (_f && !_f.done && (_b = _e.return)) _b.call(_e);
                        }
                        finally { if (e_30) throw e_30.error; }
                    }
                }
                PageInfo.runIndex();
                this.optTotalForMedium(preInfo, this.md1TotalInfo, workSheet, itemTotalFormat);
                PageInfo.runIndex();
            }
        }
        catch (e_29_1) { e_29 = { error: e_29_1 }; }
        finally {
            try {
                if (_d && !_d.done && (_a = _c.return)) _a.call(_c);
            }
            finally { if (e_29) throw e_29.error; }
        }
        PageInfo.runIndex();
        var cells = this.getCellsByFormat(workSheet, totalFormat, PageInfo);
        cells[0].value = "申告媒体①合計";
        cells[1].merge(cells[0]);
        cells[2].merge(cells[0]);
        cells[3].value = this.md1TotalInfo.sum1Total;
        cells[4].merge(cells[3]);
        cells[5].merge(cells[3]);
        cells[6].value = this.md1TotalInfo.sum2Total;
        cells[7].merge(cells[6]);
        cells[8].merge(cells[6]);
        cells[9].value = this.md1TotalInfo.Total;
        cells[10].merge(cells[9]);
        cells[11].merge(cells[9]);
        cells[9].border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" }
        };
        PageInfo.runIndex();
        PageInfo.runIndex();
    };
    AnalysReportService.prototype.sumMedium2 = function (record) {
        if (!this.optMediumMap2)
            return;
        if (record.MajorItems__c !== "インターネット")
            return;
        var allNoneFlag = true;
        for (var filedName in filedMappingMedia2) {
            if (!record[filedName])
                continue;
            allNoneFlag = false;
            var preInfo = this.optMediumMap2.get(filedName);
            if (!preInfo) {
                preInfo = new MediumInfo();
                preInfo.name = filedMappingMedia2[filedName];
                preInfo.itemList = new Map();
                this.optMediumMap2.set(filedName, preInfo);
            }
            preInfo.total++;
            this.md2TotalInfo.Total++;
            var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
            if (!responseDate)
                return;
            var setNum = this.getPeriodByDate(responseDate);
            if (!setNum)
                return;
            if (setNum === 1) {
                preInfo.countNum1++;
                this.md2TotalInfo.sum1Total++;
            }
            else {
                preInfo.countNum2++;
                this.md2TotalInfo.sum2Total++;
            }
            if (preInfo.itemList && record[filedName]) {
                var itemInfo = preInfo.itemList.get(record[filedName]);
                if (!itemInfo) {
                    itemInfo = new MediumInfo();
                    itemInfo.name = record[filedName];
                    preInfo.itemList.set(itemInfo.name, itemInfo);
                }
                if (setNum === 1) {
                    itemInfo.countNum1++;
                }
                else {
                    itemInfo.countNum2++;
                }
                itemInfo.total++;
            }
        }
        if (allNoneFlag) {
            this.md2TotalInfo.noneTotal++;
            var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
            if (!responseDate)
                return;
            var setNum = this.getPeriodByDate(responseDate);
            if (!setNum)
                return;
            if (setNum === 1) {
                this.md2TotalInfo.noneSum1++;
            }
            else {
                this.md2TotalInfo.noneSum1++;
            }
        }
    };
    AnalysReportService.prototype.outPutMedium2 = function (workSheet, formatSheet) {
        var e_31, _a, e_32, _b;
        var optMap = this.optMediumMap2;
        if (!optMap)
            return;
        var PageInfo = this.MediumPageInfo;
        PageInfo.init();
        PageInfo.colNum = 1;
        var titleFormat = formatSheet.getCell("A100");
        var itemFormat = [];
        formatSheet.getRow(101).eachCell(function (cell) {
            itemFormat.push(cell);
        });
        var itemTotalFormat = [];
        formatSheet.getRow(105).eachCell({ includeEmpty: true }, function (cell) {
            itemTotalFormat.push(cell);
        });
        var noneTotalFormat = [];
        formatSheet.getRow(107).eachCell({ includeEmpty: true }, function (cell) {
            noneTotalFormat.push(cell);
        });
        var totalFormat = [];
        formatSheet.getRow(103).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        try {
            for (var _c = tslib_1.__values(optMap.values()), _d = _c.next(); !_d.done; _d = _c.next()) {
                var preInfo = _d.value;
                var targetCell = workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex);
                targetCell.style = titleFormat.style;
                targetCell.value = preInfo.name;
                PageInfo.runIndex();
                if (preInfo.itemList && preInfo.itemList.size > 0) {
                    try {
                        for (var _e = (e_32 = void 0, tslib_1.__values(preInfo.itemList.values())), _f = _e.next(); !_f.done; _f = _e.next()) {
                            var itemInfo = _f.value;
                            var cells_17 = this.getCellsByFormat(workSheet, itemFormat, PageInfo);
                            cells_17[0].value = itemInfo.name;
                            cells_17[1].merge(cells_17[0]);
                            cells_17[2].merge(cells_17[0]);
                            cells_17[3].value = itemInfo.countNum1;
                            cells_17[4].value = preInfo.countNum1 === 0 ? 0 : itemInfo.countNum1 / preInfo.countNum1;
                            cells_17[5].value = preInfo.countNum1 === 0 ? 0 : itemInfo.countNum1 / preInfo.countNum1;
                            cells_17[6].value = itemInfo.countNum2;
                            cells_17[7].value = preInfo.countNum2 === 0 ? 0 : itemInfo.countNum2 / preInfo.countNum2;
                            cells_17[8].value = preInfo.countNum2 === 0 ? 0 : itemInfo.countNum2 / preInfo.countNum2;
                            cells_17[9].value = itemInfo.total;
                            cells_17[10].value = preInfo.total === 0 ? 0 : itemInfo.total / preInfo.total;
                            cells_17[11].value = preInfo.total === 0 ? 0 : itemInfo.total / preInfo.total;
                            PageInfo.runIndex();
                        }
                    }
                    catch (e_32_1) { e_32 = { error: e_32_1 }; }
                    finally {
                        try {
                            if (_f && !_f.done && (_b = _e.return)) _b.call(_e);
                        }
                        finally { if (e_32) throw e_32.error; }
                    }
                }
                PageInfo.runIndex();
                this.optTotalForMedium(preInfo, this.md2TotalInfo, workSheet, itemTotalFormat);
                PageInfo.runIndex();
                PageInfo.runIndex();
            }
        }
        catch (e_31_1) { e_31 = { error: e_31_1 }; }
        finally {
            try {
                if (_d && !_d.done && (_a = _c.return)) _a.call(_c);
            }
            finally { if (e_31) throw e_31.error; }
        }
        var noneCells = this.getCellsByFormat(workSheet, noneTotalFormat, PageInfo);
        noneCells[3].value = this.md2TotalInfo.noneSum1;
        noneCells[4].merge(noneCells[3]);
        noneCells[5].merge(noneCells[3]);
        noneCells[6].value = this.md2TotalInfo.noneSum2;
        noneCells[7].merge(noneCells[6]);
        noneCells[8].merge(noneCells[6]);
        noneCells[9].value = this.md2TotalInfo.noneTotal;
        noneCells[10].merge(noneCells[9]);
        noneCells[11].merge(noneCells[9]);
        noneCells[9].border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" }
        };
        PageInfo.runIndex();
        PageInfo.runIndex();
        var cells = this.getCellsByFormat(workSheet, totalFormat, PageInfo);
        cells[0].value = "申告媒体②合計";
        cells[1].merge(cells[0]);
        cells[2].merge(cells[0]);
        cells[3].value = this.md2TotalInfo.sum1Total + this.md2TotalInfo.noneSum1;
        cells[4].merge(cells[3]);
        cells[5].merge(cells[3]);
        cells[6].value = this.md2TotalInfo.sum2Total + this.md2TotalInfo.noneSum2;
        cells[7].merge(cells[6]);
        cells[8].merge(cells[6]);
        cells[9].value = this.md2TotalInfo.Total + this.md2TotalInfo.noneTotal;
        cells[10].merge(cells[9]);
        cells[11].merge(cells[9]);
        cells[9].border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" }
        };
        PageInfo.runIndex();
        PageInfo.runIndex();
    };
    AnalysReportService.prototype.sumMedium3 = function (record) {
        if (!this.optMediumMap3)
            return;
        if (record.MajorItems__c === "インターネット")
            return;
        var allNoneFlag = true;
        for (var filedName in filedMappingMedia2) {
            if (!record[filedName])
                continue;
            allNoneFlag = false;
            var preInfo = this.optMediumMap3.get(filedName);
            if (!preInfo) {
                preInfo = new MediumInfo();
                preInfo.name = filedMappingMedia2[filedName];
                preInfo.itemList = new Map();
                this.optMediumMap3.set(filedName, preInfo);
            }
            preInfo.total++;
            this.md3TotalInfo.Total++;
            var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
            if (!responseDate)
                return;
            var setNum = this.getPeriodByDate(responseDate);
            if (!setNum)
                return;
            if (setNum === 1) {
                preInfo.countNum1++;
                this.md3TotalInfo.sum1Total++;
            }
            else {
                preInfo.countNum2++;
                this.md3TotalInfo.sum2Total++;
            }
            if (preInfo.itemList && record[filedName]) {
                var itemInfo = preInfo.itemList.get(record[filedName]);
                if (!itemInfo) {
                    itemInfo = new MediumInfo();
                    itemInfo.name = record[filedName];
                    preInfo.itemList.set(itemInfo.name, itemInfo);
                }
                if (setNum === 1) {
                    itemInfo.countNum1++;
                }
                else {
                    itemInfo.countNum2++;
                }
                itemInfo.total++;
            }
        }
        if (allNoneFlag) {
            this.md3TotalInfo.noneTotal++;
            var responseDate = tools_service_1.Tools.strToDate(record.ResponseDate__c);
            if (!responseDate)
                return;
            var setNum = this.getPeriodByDate(responseDate);
            if (!setNum)
                return;
            if (setNum === 1) {
                this.md3TotalInfo.noneSum1++;
            }
            else {
                this.md3TotalInfo.noneSum1++;
            }
        }
    };
    AnalysReportService.prototype.outPutMedium3 = function (workSheet, formatSheet) {
        var e_33, _a, e_34, _b;
        var optMap = this.optMediumMap3;
        if (!optMap)
            return;
        var PageInfo = this.MediumPageInfo;
        PageInfo.init();
        PageInfo.colNum = 2;
        var titleFormat = formatSheet.getCell("A100");
        var itemFormat = [];
        formatSheet.getRow(101).eachCell(function (cell) {
            itemFormat.push(cell);
        });
        var itemTotalFormat = [];
        formatSheet.getRow(105).eachCell({ includeEmpty: true }, function (cell) {
            itemTotalFormat.push(cell);
        });
        var noneTotalFormat = [];
        formatSheet.getRow(107).eachCell({ includeEmpty: true }, function (cell) {
            noneTotalFormat.push(cell);
        });
        var totalFormat = [];
        formatSheet.getRow(103).eachCell(function (cell) {
            totalFormat.push(cell);
        });
        try {
            for (var _c = tslib_1.__values(optMap.values()), _d = _c.next(); !_d.done; _d = _c.next()) {
                var preInfo = _d.value;
                var targetCell = workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex);
                targetCell.style = titleFormat.style;
                targetCell.value = preInfo.name;
                PageInfo.runIndex();
                if (preInfo.itemList && preInfo.itemList.size > 0) {
                    try {
                        for (var _e = (e_34 = void 0, tslib_1.__values(preInfo.itemList.values())), _f = _e.next(); !_f.done; _f = _e.next()) {
                            var itemInfo = _f.value;
                            var cells_18 = this.getCellsByFormat(workSheet, itemFormat, PageInfo);
                            cells_18[0].value = itemInfo.name;
                            cells_18[1].merge(cells_18[0]);
                            cells_18[2].merge(cells_18[0]);
                            cells_18[3].value = itemInfo.countNum1;
                            cells_18[4].value = preInfo.countNum1 === 0 ? 0 : itemInfo.countNum1 / preInfo.countNum1;
                            cells_18[5].value = preInfo.countNum1 === 0 ? 0 : itemInfo.countNum1 / preInfo.countNum1;
                            cells_18[6].value = itemInfo.countNum2;
                            cells_18[7].value = preInfo.countNum2 === 0 ? 0 : itemInfo.countNum2 / preInfo.countNum2;
                            cells_18[8].value = preInfo.countNum2 === 0 ? 0 : itemInfo.countNum2 / preInfo.countNum2;
                            cells_18[9].value = itemInfo.total;
                            cells_18[10].value = preInfo.total === 0 ? 0 : itemInfo.total / preInfo.total;
                            cells_18[11].value = preInfo.total === 0 ? 0 : itemInfo.total / preInfo.total;
                            PageInfo.runIndex();
                        }
                    }
                    catch (e_34_1) { e_34 = { error: e_34_1 }; }
                    finally {
                        try {
                            if (_f && !_f.done && (_b = _e.return)) _b.call(_e);
                        }
                        finally { if (e_34) throw e_34.error; }
                    }
                }
                PageInfo.runIndex();
                this.optTotalForMedium(preInfo, this.md3TotalInfo, workSheet, itemTotalFormat);
                PageInfo.runIndex();
                PageInfo.runIndex();
            }
        }
        catch (e_33_1) { e_33 = { error: e_33_1 }; }
        finally {
            try {
                if (_d && !_d.done && (_a = _c.return)) _a.call(_c);
            }
            finally { if (e_33) throw e_33.error; }
        }
        var noneCells = this.getCellsByFormat(workSheet, noneTotalFormat, PageInfo);
        noneCells[3].value = this.md3TotalInfo.noneSum1;
        noneCells[4].merge(noneCells[3]);
        noneCells[5].merge(noneCells[3]);
        noneCells[6].value = this.md3TotalInfo.noneSum2;
        noneCells[7].merge(noneCells[6]);
        noneCells[8].merge(noneCells[6]);
        noneCells[9].value = this.md3TotalInfo.noneTotal;
        noneCells[10].merge(noneCells[9]);
        noneCells[11].merge(noneCells[9]);
        noneCells[9].border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" }
        };
        PageInfo.runIndex();
        PageInfo.runIndex();
        var cells = this.getCellsByFormat(workSheet, totalFormat, PageInfo);
        cells[0].value = "申告媒体②合計";
        cells[1].merge(cells[0]);
        cells[2].merge(cells[0]);
        cells[3].value = this.md3TotalInfo.sum1Total + this.md3TotalInfo.noneSum1;
        cells[4].merge(cells[3]);
        cells[5].merge(cells[3]);
        cells[6].value = this.md3TotalInfo.sum2Total + this.md3TotalInfo.noneSum2;
        cells[7].merge(cells[6]);
        cells[8].merge(cells[6]);
        cells[9].value = this.md3TotalInfo.Total + this.md3TotalInfo.noneTotal;
        cells[10].merge(cells[9]);
        cells[11].merge(cells[9]);
        cells[9].border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" }
        };
        PageInfo.runIndex();
        PageInfo.runIndex();
    };
    AnalysReportService.prototype.setOtherAnalysis = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var questionRes, questionMap, _a, _b, question, answerRes, answerMap, _c, _d, answer, answerData, questionNo, question, answerInfo, answerItem, subOption;
            var e_35, _e, e_36, _f;
            return tslib_1.__generator(this, function (_g) {
                switch (_g.label) {
                    case 0:
                        if (!this.analysisRptConfig.subSurvey)
                            return [2];
                        return [4, this.sfdcService.query("SELECT Id,Name,QuestionNumber__c,QuestionContent__c,(SELECT OptionItemNumber__c,OptionItemValue__c FROM optionProjectQuestionnaireQuestions__r) from ProjectQuestionnaireQuestions__c WHERE QuestionnaireAccount__c = '" + this.analysisRptConfig.subSurvey + "' AND QuestionDeleteFlag__c = false")];
                    case 1:
                        questionRes = _g.sent();
                        if (questionRes.totalSize === 0)
                            return [2];
                        questionMap = new Map();
                        try {
                            for (_a = tslib_1.__values(questionRes.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                question = _b.value;
                                if (!question.QuestionNumber__c)
                                    continue;
                                questionMap.set(question.QuestionNumber__c, question);
                            }
                        }
                        catch (e_35_1) { e_35 = { error: e_35_1 }; }
                        finally {
                            try {
                                if (_b && !_b.done && (_e = _a.return)) _e.call(_a);
                            }
                            finally { if (e_35) throw e_35.error; }
                        }
                        return [4, this.sfdcService.query("SELECT Id,AnswersContents__c,Opportunity__r.ResponseDate__c FROM QuestionnaireAnswer__c WHERE ProjectQuestionnaire__c = '" + this.analysisRptConfig.subSurvey + "'")];
                    case 2:
                        answerRes = _g.sent();
                        if (answerRes.totalSize === 0)
                            return [2];
                        answerMap = new Map();
                        this.optAnswerMap = answerMap;
                        try {
                            for (_c = tslib_1.__values(answerRes.records), _d = _c.next(); !_d.done; _d = _c.next()) {
                                answer = _d.value;
                                if (!answer.AnswersContents__c)
                                    continue;
                                answerData = JSON.parse(answer.AnswersContents__c);
                                for (questionNo in answerData) {
                                    question = questionMap.get(questionNo);
                                    if (!question)
                                        continue;
                                    answerInfo = answerMap.get(question.QuestionContent__c);
                                    if (!answerInfo) {
                                        answerInfo = new AnswerInfo();
                                        answerInfo.name = question.QuestionContent__c;
                                        answerMap.set(answerInfo.name, answerInfo);
                                    }
                                    answerItem = void 0;
                                    if (typeof answerData[questionNo] === "string") {
                                        if (!answerData[questionNo])
                                            continue;
                                        answerItem = answerInfo.itemList.get(answerData[questionNo]);
                                        if (!answerItem) {
                                            answerItem = new AnswerInfo();
                                            answerItem.name = answerData[questionNo];
                                            answerInfo.itemList.set(answerItem.name, answerItem);
                                            this.sumAnswer(answerItem, answerInfo, answer.Opportunity__r.ResponseDate__c);
                                        }
                                    }
                                    else {
                                        for (subOption in answerData[questionNo]) {
                                            answerItem = answerInfo.itemList.get(subOption);
                                            if (!answerItem) {
                                                answerItem = new AnswerInfo();
                                                answerItem.name = subOption;
                                                answerInfo.itemList.set(answerItem.name, answerItem);
                                            }
                                            this.sumAnswer(answerItem, answerInfo, answer.Opportunity__r.ResponseDate__c);
                                        }
                                    }
                                }
                            }
                        }
                        catch (e_36_1) { e_36 = { error: e_36_1 }; }
                        finally {
                            try {
                                if (_d && !_d.done && (_f = _c.return)) _f.call(_c);
                            }
                            finally { if (e_36) throw e_36.error; }
                        }
                        return [2];
                }
            });
        });
    };
    AnalysReportService.prototype.sumAnswer = function (answerItem, question, ResponseDate__c) {
        answerItem.total++;
        question.total++;
        var responseDate = tools_service_1.Tools.strToDate(ResponseDate__c);
        if (!responseDate)
            return;
        var setNum = this.getPeriodByDate(responseDate);
        if (!setNum)
            return;
        if (setNum === 1) {
            question.countNum1++;
            answerItem.countNum1++;
        }
        else {
            question.countNum2++;
            answerItem.countNum2++;
        }
    };
    AnalysReportService.prototype.outPutOther = function (workSheet, formatSheet) {
        var e_37, _a, e_38, _b;
        var optMap = this.optAnswerMap;
        if (!optMap)
            return;
        var PageInfo = this.OtherPageInfo;
        var itemFormat = [];
        formatSheet.getRow(121).eachCell(function (cell) {
            itemFormat.push(cell);
        });
        var itemTotalFormat = [];
        formatSheet.getRow(123).eachCell({ includeEmpty: true }, function (cell) {
            itemTotalFormat.push(cell);
        });
        try {
            for (var _c = tslib_1.__values(optMap.values()), _d = _c.next(); !_d.done; _d = _c.next()) {
                var preInfo = _d.value;
                this.setOtherTitle(workSheet, formatSheet, 116, preInfo.name);
                try {
                    for (var _e = (e_38 = void 0, tslib_1.__values(preInfo.itemList.values())), _f = _e.next(); !_f.done; _f = _e.next()) {
                        var itemInfo = _f.value;
                        var cells = this.getCellsByFormat(workSheet, itemFormat, PageInfo);
                        cells[0].value = itemInfo.name;
                        cells[1].merge(cells[0]);
                        cells[2].value = itemInfo.countNum1;
                        cells[3].value = preInfo.countNum1 === 0 ? 0 : itemInfo.countNum1 / preInfo.countNum1;
                        cells[4].value = itemInfo.countNum2;
                        cells[5].value = preInfo.countNum2 === 0 ? 0 : itemInfo.countNum2 / preInfo.countNum2;
                        cells[6].value = itemInfo.total;
                        cells[7].value = preInfo.total === 0 ? 0 : itemInfo.total / preInfo.total;
                        PageInfo.runIndex();
                    }
                }
                catch (e_38_1) { e_38 = { error: e_38_1 }; }
                finally {
                    try {
                        if (_f && !_f.done && (_b = _e.return)) _b.call(_e);
                    }
                    finally { if (e_38) throw e_38.error; }
                }
                PageInfo.runIndex();
                var itemTotalCells = this.getCellsByFormat(workSheet, itemTotalFormat, PageInfo);
                itemTotalCells[0].value = "合計";
                itemTotalCells[2].value = preInfo.countNum1;
                itemTotalCells[3].merge(itemTotalCells[2]);
                itemTotalCells[4].value = preInfo.countNum2;
                itemTotalCells[5].merge(itemTotalCells[4]);
                itemTotalCells[6].value = preInfo.total;
                itemTotalCells[7].merge(itemTotalCells[6]);
                itemTotalCells[6].border = {
                    top: { style: "thin" },
                    left: { style: "thin" },
                    bottom: { style: "thin" },
                    right: { style: "thin" }
                };
                PageInfo.runIndex();
                PageInfo.runIndex();
            }
        }
        catch (e_37_1) { e_37 = { error: e_37_1 }; }
        finally {
            try {
                if (_d && !_d.done && (_a = _c.return)) _a.call(_c);
            }
            finally { if (e_37) throw e_37.error; }
        }
    };
    AnalysReportService.prototype.getWorkSheetByName = function (sheetName) {
        return this.sheetLsit.find(function (sheet) {
            return sheet.name === sheetName;
        });
    };
    AnalysReportService.prototype.getCellsByFormat = function (workSheet, formatCells, PageInfo) {
        if (!PageInfo) {
            PageInfo = this.PageInfo;
        }
        var cells = [];
        for (var index = 0; index < formatCells.length; index++) {
            var formatCell = formatCells[index];
            var targetCell = workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex + index);
            targetCell.style = formatCell.style;
            cells.push(targetCell);
        }
        return cells;
    };
    AnalysReportService.PAGE_SIZE = 63;
    return AnalysReportService;
}());
exports.AnalysReportService = AnalysReportService;
//# sourceMappingURL=analysisReport.service.js.map