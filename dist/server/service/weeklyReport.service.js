"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.WeeklyReportService = void 0;
var tslib_1 = require("tslib");
var exceljs_1 = tslib_1.__importDefault(require("exceljs"));
var sfdc_service_1 = require("./sfdc.service");
var tools_service_1 = require("./tools.service");
var date_and_time_1 = tslib_1.__importDefault(require("date-and-time"));
var pg_service_1 = require("./pg.service");
var csvtojson_1 = tslib_1.__importDefault(require("csvtojson"));
var Footstay = (function () {
    function Footstay() {
        this.yearStart = new Date(new Date().getFullYear(), 0, 1);
        this.month3Start = new Date();
        this.month3Visit = new Set();
        this.yearVisit = new Set();
    }
    return Footstay;
}());
var BrillaClubInfo = (function () {
    function BrillaClubInfo() {
        this.visitNum = 0;
        this.joinedNum = 0;
        this.liveJoinNum = 0;
    }
    Object.defineProperty(BrillaClubInfo.prototype, "notJoinNum", {
        get: function () {
            return this.visitNum - this.joinedNum;
        },
        enumerable: false,
        configurable: true
    });
    return BrillaClubInfo;
}());
var SebQuestionInfo = (function () {
    function SebQuestionInfo() {
        this.weekPriceImpression = new Set();
        this.weekPriceImpressionCount = new Set();
        this.weekLocalImpression = new Set();
        this.weekLocalImpressionCount = new Set();
        this.weekExaminationDegree = new Set();
        this.weekExaminationDegreeCount = new Set();
        this.weekMREvaluation = new Set();
        this.weekMREvaluationCount = new Set();
        this.PriceImpression = ["やや高い", "高い"];
        this.ExaminationDegree = ["前向きに検討したい", "是非購入したい"];
        this.LocalImpression = ["まあまあ良い", "大変良い"];
        this.MREvaluation = ["まあまあ良い", "大変良い"];
    }
    return SebQuestionInfo;
}());
var SupplyInfo = (function () {
    function SupplyInfo() {
        this.supplyedCount = 0;
        this.reservationCount = 0;
        this.applicationCount = 0;
        this.contractCount = 0;
        this.noneCount = 0;
        this.deliverCount = 0;
        this.rowIndex = 0;
    }
    return SupplyInfo;
}());
var WeeklyReportService = (function () {
    function WeeklyReportService() {
        this.noneInfo = {
            count: 0
        };
        this.workbook = new exceljs_1.default.Workbook();
        this.sfdcService = new sfdc_service_1.SfdcService();
        this.pgService = new pg_service_1.PgService();
    }
    WeeklyReportService.prototype.getFileBuffer = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4, this.workbook.xlsx.writeBuffer()];
                    case 1: return [2, _a.sent()];
                }
            });
        });
    };
    WeeklyReportService.prototype.gennaretaWeeklyReport = function (weeklyRptConfig) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                this.excute(weeklyRptConfig).catch(function (err) {
                    throw err;
                });
                return [2];
            });
        });
    };
    WeeklyReportService.prototype.excute = function (weeklyRptConfig) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var _a, fromDate, toDate, queryStartDate, queryEndQuery, _b, _c, item, promiseJobList;
            var e_1, _d;
            return tslib_1.__generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        _a = this;
                        return [4, this.workbook.xlsx.readFile("server/template/週報テンプレート.xlsx")];
                    case 1:
                        _a.workbook = _e.sent();
                        this.worksheet = this.workbook.worksheets[0];
                        try {
                            for (_b = tslib_1.__values(weeklyRptConfig.weeklyList), _c = _b.next(); !_c.done; _c = _b.next()) {
                                item = _c.value;
                                item.PeriodFrom = tools_service_1.Tools.strToDate(item.PeriodFrom__c);
                                item.PeriodTo = tools_service_1.Tools.strToDate(item.PeriodTo__c);
                                if (item.WeekName !== "次週") {
                                    if (!fromDate || (item.PeriodFrom && fromDate > item.PeriodFrom)) {
                                        fromDate = item.PeriodFrom;
                                        queryStartDate = item.PeriodFrom__c;
                                    }
                                    if (!toDate || (item.PeriodTo && toDate < item.PeriodTo)) {
                                        toDate = item.PeriodTo;
                                        queryEndQuery = item.PeriodTo__c;
                                    }
                                }
                                if (item.WeekName === "先々週") {
                                    this.beforeLastWeek = item;
                                }
                                else if (item.WeekName === "先週") {
                                    this.lastWeek = item;
                                }
                                else if (item.WeekName === "今週") {
                                    this.theWeek = item;
                                }
                                else if (item.WeekName === "次週") {
                                    this.nextWeek = item;
                                }
                            }
                        }
                        catch (e_1_1) { e_1 = { error: e_1_1 }; }
                        finally {
                            try {
                                if (_c && !_c.done && (_d = _b.return)) _d.call(_b);
                            }
                            finally { if (e_1) throw e_1.error; }
                        }
                        this.initFootStay();
                        this.initMonthlyCustomerTotalGoal(weeklyRptConfig);
                        this.initSebQuestionInfo();
                        promiseJobList = [];
                        promiseJobList.push(this.setTitleAndAvPrice(weeklyRptConfig));
                        promiseJobList.push(this.setSalesPeriodMaster(weeklyRptConfig));
                        promiseJobList.push(this.setWeeklyInfo(weeklyRptConfig, queryStartDate, queryEndQuery));
                        promiseJobList.push(this.setBrilliaClub(weeklyRptConfig));
                        promiseJobList.push(this.setOOSInfo(weeklyRptConfig));
                        promiseJobList.push(this.saveContractGoal(weeklyRptConfig));
                        promiseJobList.push(this.setSupplystatus(weeklyRptConfig));
                        return [4, Promise.all(promiseJobList).catch(function (err) {
                                throw err;
                            })];
                    case 2:
                        _e.sent();
                        this.saveMonthlyCustomerTotalGoal();
                        this.saveSubQuestion();
                        this.saveOtherData();
                        return [2];
                }
            });
        });
    };
    WeeklyReportService.prototype.setTitleAndAvPrice = function (weeklyRptConfig) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var accountRes, acc, count, _a, _b, build, val, _c, _d, sellerInfo, _e, _f, saleInfo, buildIdList, _g, _h, build, buildIdSql, roomRes, avInfo, buildAvMap, _j, _k, room, buildAvInfo, totalPrice, bulidSellText, bulidFormalText, bulidPlayText, avPrice, _l, _m, buildId, avPriceInfo;
            var e_2, _o, e_3, _p, e_4, _q, e_5, _r, e_6, _s, e_7, _t;
            return tslib_1.__generator(this, function (_u) {
                switch (_u.label) {
                    case 0:
                        if (!this.worksheet)
                            return [2];
                        return [4, this.sfdcService
                                .query("SELECT Id,Name,LocationLotNumber__c,(SELECT Id,OldeBrilliaTran__c FROM Buildings__r),(SELECT Id,CorporationName__c,Share__c FROM Project_SellerInfo__r order by SortOrder__c),(SELECT Id,CorporationName__c,Share__c FROM Project_SaleInfo__r order by SortOrder__c) FROM Account WHERE Id = '" + weeklyRptConfig.projectId + "'")
                                .catch(function (err) {
                                throw err;
                            })];
                    case 1:
                        accountRes = _u.sent();
                        acc = accountRes.records[0];
                        this.worksheet.getCell("D2").value = acc.Name;
                        this.worksheet.getCell("C3").value =
                            (weeklyRptConfig.selectWeek.Kinds__c === "問合せ" ? "問合せ" : "") + weeklyRptConfig.selectWeek.Weeks__c;
                        this.worksheet.getCell("C6").value = acc.LocationLotNumber__c;
                        count = 0;
                        if (acc.Buildings__r && acc.Buildings__r.totalSize > 0) {
                            try {
                                for (_a = tslib_1.__values(acc.Buildings__r.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                    build = _b.value;
                                    if (build.OldeBrilliaTran__c && count < 2) {
                                        count++;
                                        this.worksheet.getCell("C8").value = build.OldeBrilliaTran__c;
                                        break;
                                    }
                                }
                            }
                            catch (e_2_1) { e_2 = { error: e_2_1 }; }
                            finally {
                                try {
                                    if (_b && !_b.done && (_o = _a.return)) _o.call(_a);
                                }
                                finally { if (e_2) throw e_2.error; }
                            }
                        }
                        val = "";
                        if (acc.Project_SellerInfo__r && acc.Project_SellerInfo__r.totalSize > 0) {
                            try {
                                for (_c = tslib_1.__values(acc.Project_SellerInfo__r.records), _d = _c.next(); !_d.done; _d = _c.next()) {
                                    sellerInfo = _d.value;
                                    if (sellerInfo.CorporationName__c) {
                                        val += sellerInfo.CorporationName__c + "(" + sellerInfo.Share__c + ")\u3001";
                                    }
                                }
                            }
                            catch (e_3_1) { e_3 = { error: e_3_1 }; }
                            finally {
                                try {
                                    if (_d && !_d.done && (_p = _c.return)) _p.call(_c);
                                }
                                finally { if (e_3) throw e_3.error; }
                            }
                            val = val.length > 0 ? val.substring(0, val.length - 1) : "";
                            this.worksheet.getCell("C11").value = val;
                        }
                        val = "";
                        if (acc.Project_SaleInfo__r && acc.Project_SaleInfo__r.totalSize > 0) {
                            try {
                                for (_e = tslib_1.__values(acc.Project_SaleInfo__r.records), _f = _e.next(); !_f.done; _f = _e.next()) {
                                    saleInfo = _f.value;
                                    if (saleInfo.CorporationName__c) {
                                        val += saleInfo.CorporationName__c + "(" + saleInfo.Share__c + ")\u3001";
                                    }
                                }
                            }
                            catch (e_4_1) { e_4 = { error: e_4_1 }; }
                            finally {
                                try {
                                    if (_f && !_f.done && (_q = _e.return)) _q.call(_e);
                                }
                                finally { if (e_4) throw e_4.error; }
                            }
                            val = val.length > 0 ? val.substring(0, val.length - 1) : "";
                            this.worksheet.getCell("C13").value = val;
                        }
                        this.worksheet.getCell("O3").value = weeklyRptConfig.questionStart
                            ? weeklyRptConfig.questionStart.replace(/-/g, "/")
                            : "";
                        this.worksheet.getCell("S3").value = weeklyRptConfig.mrOpen ? weeklyRptConfig.mrOpen.replace(/-/g, "/") : "";
                        this.worksheet.getCell("W3").value = weeklyRptConfig.extraditeStart;
                        if (!acc.Buildings__r) return [3, 3];
                        buildIdList = [];
                        try {
                            for (_g = tslib_1.__values(acc.Buildings__r.records), _h = _g.next(); !_h.done; _h = _g.next()) {
                                build = _h.value;
                                buildIdList.push(build.Id);
                            }
                        }
                        catch (e_5_1) { e_5 = { error: e_5_1 }; }
                        finally {
                            try {
                                if (_h && !_h.done && (_r = _g.return)) _r.call(_g);
                            }
                            finally { if (e_5) throw e_5.error; }
                        }
                        buildIdSql = sfdc_service_1.SfdcService.getInSql(buildIdList);
                        if (!(buildIdList.length > 0)) return [3, 3];
                        return [4, this.sfdcService
                                .query("SELECT Id,Name,SellingPrice__c,PrePriceWithTax__c,FullPriceFlag__c,ForAndnonForSaleStore__c,Building__c,BuildingName__c,OccupiedArea__c FROM Room__c WHERE Building__c IN " + buildIdSql)
                                .catch(function (err) {
                                throw err;
                            })];
                    case 2:
                        roomRes = _u.sent();
                        if (roomRes.totalSize > 0) {
                            avInfo = {
                                sellNum: 0,
                                totalNum: roomRes.totalSize,
                                formalNum: 0,
                                playNum: 0,
                                sellTotalPrice: 0.0,
                                sellArea: 0,
                                formalTotalPrice: 0.0,
                                formalArea: 0,
                                playTotalPrice: 0.0,
                                playArea: 0
                            };
                            buildAvMap = new Map();
                            try {
                                for (_j = tslib_1.__values(roomRes.records), _k = _j.next(); !_k.done; _k = _j.next()) {
                                    room = _k.value;
                                    buildAvInfo = buildAvMap.get(room.Building__c);
                                    if (!buildAvInfo) {
                                        buildAvInfo = {
                                            buildingName: room.BuildingName__c,
                                            sellNum: 0,
                                            totalNum: roomRes.totalSize,
                                            formalNum: 0,
                                            playNum: 0,
                                            sellTotalPrice: 0.0,
                                            sellArea: 0,
                                            formalTotalPrice: 0.0,
                                            formalArea: 0,
                                            playTotalPrice: 0.0,
                                            playArea: 0
                                        };
                                        buildAvMap.set(room.Building__c, buildAvInfo);
                                    }
                                    if (room.ForAndnonForSaleStore__c === "分譲（一般）") {
                                        avInfo.sellNum++;
                                        totalPrice = 0;
                                        if (room.FullPriceFlag__c === "済") {
                                            totalPrice = room.SellingPrice__c;
                                        }
                                        else {
                                            totalPrice = room.PrePriceWithTax__c;
                                        }
                                        avInfo.sellTotalPrice += totalPrice ? totalPrice : 0;
                                        avInfo.sellArea += room.OccupiedArea__c ? room.OccupiedArea__c : 0;
                                        buildAvInfo.sellNum++;
                                        buildAvInfo.sellTotalPrice += totalPrice ? totalPrice : 0;
                                        buildAvInfo.sellArea += room.OccupiedArea__c ? room.OccupiedArea__c : 0;
                                    }
                                    if (room.FullPriceFlag__c === "済") {
                                        avInfo.formalNum++;
                                        avInfo.formalTotalPrice += room.SellingPrice__c ? room.SellingPrice__c : 0;
                                        avInfo.formalArea += room.OccupiedArea__c ? room.OccupiedArea__c : 0;
                                        buildAvInfo.formalNum++;
                                        buildAvInfo.formalTotalPrice += room.SellingPrice__c ? room.SellingPrice__c : 0;
                                        buildAvInfo.formalArea += room.OccupiedArea__c ? room.OccupiedArea__c : 0;
                                    }
                                    else {
                                        avInfo.playNum++;
                                        avInfo.playTotalPrice += room.PrePriceWithTax__c ? room.PrePriceWithTax__c : 0;
                                        avInfo.playArea += room.OccupiedArea__c ? room.OccupiedArea__c : 0;
                                        buildAvInfo.playNum++;
                                        buildAvInfo.playTotalPrice += room.PrePriceWithTax__c ? room.PrePriceWithTax__c : 0;
                                        buildAvInfo.playArea += room.OccupiedArea__c ? room.OccupiedArea__c : 0;
                                    }
                                }
                            }
                            catch (e_6_1) { e_6 = { error: e_6_1 }; }
                            finally {
                                try {
                                    if (_k && !_k.done && (_s = _j.return)) _s.call(_j);
                                }
                                finally { if (e_6) throw e_6.error; }
                            }
                            this.worksheet.getCell("P6").value = avInfo.sellNum + "\u6238(" + avInfo.totalNum + "\u6238)";
                            this.worksheet.getCell("P9").value = avInfo.formalNum;
                            this.worksheet.getCell("P12").value = avInfo.playNum;
                            console.log("sellTotalPrice:" + avInfo.sellTotalPrice);
                            console.log("sellArea:" + avInfo.sellArea);
                            this.worksheet.getCell("R6").value =
                                "@" + (avInfo.sellTotalPrice / avInfo.sellArea / 0.3025 / 10000).toFixed(1);
                            console.log("playTotalPrice:" + avInfo.formalTotalPrice);
                            console.log("playArea:" + avInfo.formalArea);
                            this.worksheet.getCell("R9").value =
                                "@" + (avInfo.formalTotalPrice / avInfo.formalArea / 0.3025 / 10000).toFixed(1);
                            console.log("playTotalPrice:" + avInfo.playTotalPrice);
                            console.log("playArea:" + avInfo.playArea);
                            this.worksheet.getCell("R12").value =
                                "@" + (avInfo.playTotalPrice / avInfo.playArea / 0.3025 / 10000).toFixed(1);
                            bulidSellText = "";
                            bulidFormalText = "";
                            bulidPlayText = "";
                            avPrice = void 0;
                            try {
                                for (_l = tslib_1.__values(buildAvMap.keys()), _m = _l.next(); !_m.done; _m = _l.next()) {
                                    buildId = _m.value;
                                    avPriceInfo = buildAvMap.get(buildId);
                                    if (!avPriceInfo)
                                        continue;
                                    avPrice = (avPriceInfo.sellTotalPrice / avPriceInfo.sellArea / 0.3025 / 10000).toFixed(1);
                                    bulidSellText += "" + avPriceInfo.buildingName + avPriceInfo.sellNum + "\u6238(" + avPriceInfo.totalNum + "\u6238)@" + avPrice + "\u3001";
                                    avPrice = (avPriceInfo.formalTotalPrice / avPriceInfo.formalArea / 0.3025 / 10000).toFixed(1);
                                    bulidFormalText += "" + avPriceInfo.buildingName + avPriceInfo.formalNum + "\u6238@" + avPrice + "\u3001";
                                    avPrice = (avPriceInfo.playTotalPrice / avPriceInfo.playArea / 0.3025 / 10000).toFixed(1);
                                    bulidPlayText += "" + avPriceInfo.buildingName + avPriceInfo.playNum + "\u6238@" + avPrice + "\u3001";
                                }
                            }
                            catch (e_7_1) { e_7 = { error: e_7_1 }; }
                            finally {
                                try {
                                    if (_m && !_m.done && (_t = _l.return)) _t.call(_l);
                                }
                                finally { if (e_7) throw e_7.error; }
                            }
                            this.worksheet.getCell("U6").value = bulidSellText.substring(0, bulidSellText.length - 1);
                            this.worksheet.getCell("U9").value = bulidFormalText.substring(0, bulidFormalText.length - 1);
                            this.worksheet.getCell("U12").value = bulidPlayText.substring(0, bulidPlayText.length - 1);
                        }
                        _u.label = 3;
                    case 3: return [2];
                }
            });
        });
    };
    WeeklyReportService.prototype.initWeekDate = function (theDate, xindex, weekName) {
        return {
            date: theDate,
            xindex: xindex,
            weekName: weekName,
            InquiryDate__c: new Set(),
            OSSRelationLatestDate__c: new Set(),
            VisitDate__c: new Set(),
            ReturnVisitDate__c: new Set(),
            ReturnLatestVisitDate__c: new Set(),
            RegisterDate__c: new Set(),
            RegisterCancelDate__c: new Set(),
            InquiryVisitNum: new Set(),
            ReserveDate__c: 0,
            ReserveCancell: 0,
            ApplicationDate__c: 0,
            ApplicationCancell: 0,
            ContractDate__c: 0,
            ContractCancell: 0,
            DeliverDate__c: 0,
            NoneNum: 0,
            NoneCancell: 0
        };
    };
    WeeklyReportService.prototype.initFootStay = function () {
        if (!this.theWeek || !this.theWeek.PeriodFrom)
            return;
        this.footstay = new Footstay();
        this.footstay.month3Start = date_and_time_1.default.addMonths(this.theWeek.PeriodFrom, -3);
        this.footstay.yearStart = new Date(this.theWeek.PeriodFrom.getFullYear(), 0, 1);
    };
    WeeklyReportService.prototype.sumFootstay = function (record) {
        if (!this.footstay)
            return;
        var visitDate = tools_service_1.Tools.strToDate(record.VisitDate__c);
        if (visitDate) {
            if (visitDate >= this.footstay.month3Start) {
                this.footstay.month3Visit.add(record.Opportunity__c + "_" + record.VisitDate__c);
            }
            if (visitDate >= this.footstay.yearStart) {
                this.footstay.yearVisit.add(record.Opportunity__c + "_" + record.VisitDate__c);
            }
        }
    };
    WeeklyReportService.prototype.setWeeklyInfo = function (weeklyRptConfig, queryStartDate, queryEndDate) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var weekInfo, startDate, endDate, weekDateMap, index, theDate, index, theDate, recordStream, csvToJsonParser, readStream, recordsTotal;
            var _this = this;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.worksheet)
                            return [2];
                        weekDateMap = new Map();
                        if (this.beforeLastWeek) {
                            weekDateMap.set("先々週", this.initWeekDate(undefined, 3, "先々週-週計"));
                        }
                        if (this.lastWeek) {
                            weekInfo = this.lastWeek;
                            startDate = weekInfo.PeriodFrom ? weekInfo.PeriodFrom.getMonth() + 1 + "/" + weekInfo.PeriodFrom.getDate() : "";
                            endDate = weekInfo.PeriodTo ? weekInfo.PeriodTo.getMonth() + 1 + "/" + weekInfo.PeriodTo.getDate() : "";
                            if (weekInfo.PeriodFrom) {
                                for (index = 0; index < 7; index++) {
                                    theDate = date_and_time_1.default.addDays(weekInfo.PeriodFrom, index);
                                    this.worksheet.getCell(32, index * 2 + 4).value = date_and_time_1.default.format(theDate, "M/DD");
                                    weekDateMap.set(date_and_time_1.default.format(theDate, "YYYY-MM-DD"), this.initWeekDate(theDate, index * 2 + 4, "先週"));
                                }
                                weekDateMap.set("先週", this.initWeekDate(undefined, 18, "先週-週計"));
                            }
                            this.worksheet.getCell("D16").value = weekInfo.WeekName + "(" + weekInfo.Weeks__c + " " + startDate + "~" + endDate + ")";
                            this.worksheet.getCell("D31").value = weekInfo.WeekName + "(" + weekInfo.Weeks__c + " " + startDate + "~" + endDate + ")";
                            this.worksheet.getCell("K29").value = weekInfo.NewGoalWeek__c + "/" + weekInfo.NewGoalWeekTotal__c + "\u4EF6";
                            this.worksheet.getCell("S29").value = weekInfo.ReturnGoalWeek__c + "/" + weekInfo.ReturnGoalWeekTotal__c + "\u4EF6";
                        }
                        if (this.theWeek) {
                            weekInfo = this.theWeek;
                            startDate = weekInfo.PeriodFrom ? weekInfo.PeriodFrom.getMonth() + 1 + "/" + weekInfo.PeriodFrom.getDate() : "";
                            endDate = weekInfo.PeriodTo ? weekInfo.PeriodTo.getMonth() + 1 + "/" + weekInfo.PeriodTo.getDate() : "";
                            if (weekInfo.PeriodFrom) {
                                for (index = 0; index < 7; index++) {
                                    theDate = date_and_time_1.default.addDays(weekInfo.PeriodFrom, index);
                                    this.worksheet.getCell(32, index * 2 + 20).value = date_and_time_1.default.format(theDate, "M/DD");
                                    weekDateMap.set(date_and_time_1.default.format(theDate, "YYYY-MM-DD"), this.initWeekDate(theDate, index * 2 + 20, "今週"));
                                }
                                weekDateMap.set("今週", this.initWeekDate(undefined, 34, "今週-週計"));
                            }
                            this.worksheet.getCell("T16").value = weekInfo.WeekName + "(" + weekInfo.Weeks__c + " " + startDate + "~" + endDate + ")";
                            this.worksheet.getCell("T31").value = weekInfo.WeekName + "(" + weekInfo.Weeks__c + " " + startDate + "~" + endDate + ")";
                            this.worksheet.getCell("AA29").value = weekInfo.NewGoalWeek__c + "/" + weekInfo.NewGoalWeekTotal__c + "\u4EF6";
                            this.worksheet.getCell("AI29").value = weekInfo.ReturnGoalWeek__c + "/" + weekInfo.ReturnGoalWeekTotal__c + "\u4EF6";
                        }
                        if (this.nextWeek) {
                            weekInfo = this.nextWeek;
                            startDate = weekInfo.PeriodFrom ? weekInfo.PeriodFrom.getMonth() + 1 + "/" + weekInfo.PeriodFrom.getDate() : "";
                            endDate = weekInfo.PeriodTo ? weekInfo.PeriodTo.getMonth() + 1 + "/" + weekInfo.PeriodTo.getDate() : "";
                            this.worksheet.getCell("AJ16").value = weekInfo.WeekName + "(" + weekInfo.Weeks__c + " " + startDate + "~" + endDate + ")";
                        }
                        weekDateMap.set("累計", this.initWeekDate(undefined, 35, "累計"));
                        return [4, this.sfdcService
                                .bulkQuery("SELECT Id,RecordTypeId,Name,Opportunity__c,Opportunity__r.OpportunityHistory__c,Opportunity__r.InquiryDate__c,Opportunity__r.FirstReturnVisitDate__c,InquiryDate__c,OSSRelationLatestDate__c,VisitDate__c,ReserveDate__c,ReturnVisitDate__c,ReturnLatestVisitDate__c,RegisterDate__c,RegisterCancelDate__c,LocalImpression__c,MREvaluation__c,PriceImpression__c,ExaminationDegree__c FROM OpportunityHistory__c WHERE Opportunity__r.AccountId = '" + weeklyRptConfig.projectId + "' AND (InquiryDate__c <= " + queryEndDate + " OR OSSRelationLatestDate__c <= " + queryEndDate + " OR VisitDate__c <= " + queryEndDate + " OR ReturnVisitDate__c <= " + queryEndDate + " OR ReturnLatestVisitDate__c <= " + queryEndDate + " OR RegisterDate__c <= " + queryEndDate + " OR RegisterCancelDate__c <= " + queryEndDate + " OR ReserveDate__c <= " + queryEndDate + ")")
                                .catch(function (err) {
                                throw err;
                            })];
                    case 1:
                        recordStream = _a.sent();
                        csvToJsonParser = csvtojson_1.default({ checkType: true });
                        readStream = recordStream.stream();
                        readStream.pipe(csvToJsonParser);
                        recordsTotal = 0;
                        csvToJsonParser.on("data", function (data) {
                            recordsTotal++;
                            var record = JSON.parse(data.toString("utf8"));
                            _this.sumTableInfo(record, weekDateMap);
                            _this.sumFootstay(record);
                            _this.setMonthlyCustomerTotalGoal(record);
                            _this.sumSubQuestion(record);
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
                            }).then(function (record) {
                                console.log("sucees record total:" + recordsTotal);
                            })];
                    case 2:
                        _a.sent();
                        return [4, this.saveTableInfo(weeklyRptConfig, weekDateMap, queryStartDate, queryEndDate)];
                    case 3:
                        _a.sent();
                        return [2];
                }
            });
        });
    };
    WeeklyReportService.prototype.setSalesPeriodMaster = function (weeklyRptConfig) {
        var _a, _b, _c, _d;
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var query2, salesPeriodMasterId, salesPeriodMap, result2, _e, _f, aggRes, spTgt, spTgt, query, result, optList, salesPeriodGroupIdList, _g, _h, spm, spTgt, spTgt, rowIndex, optList_1, optList_1_1, spInfo;
            var e_8, _j, e_9, _k, e_10, _l;
            return tslib_1.__generator(this, function (_m) {
                switch (_m.label) {
                    case 0:
                        if (!this.worksheet)
                            return [2];
                        query2 = "SELECT SalesPeriodAndStage__r.SalesPeriodGroupMaster__c SalesPeriodGroupMaster__c,SalesPeriodAndStage__c,count(Id) RoomTotal FROM Room__c WHERE SalesPeriodAndStage__c != null AND Project__c = '" + weeklyRptConfig.projectId + "' GROUP BY SalesPeriodAndStage__r.SalesPeriodGroupMaster__c,SalesPeriodAndStage__c";
                        salesPeriodMasterId = [];
                        salesPeriodMap = new Map();
                        return [4, this.sfdcService.query(query2).catch(function (err) {
                                throw err;
                            })];
                    case 1:
                        result2 = _m.sent();
                        if (result2.totalSize === 0)
                            return [2];
                        try {
                            for (_e = tslib_1.__values(result2.records), _f = _e.next(); !_f.done; _f = _e.next()) {
                                aggRes = _f.value;
                                salesPeriodMasterId.push(aggRes.SalesPeriodAndStage__c);
                                if (aggRes.SalesPeriodGroupMaster__c) {
                                    spTgt = salesPeriodMap.get(aggRes.SalesPeriodGroupMaster__c);
                                    if (spTgt) {
                                        if (spTgt)
                                            spTgt.RoomTotal += Number(aggRes.RoomTotal);
                                    }
                                    else {
                                        spTgt = {
                                            Id: "",
                                            Name: "",
                                            RoomTotal: Number(aggRes.RoomTotal)
                                        };
                                        salesPeriodMap.set(aggRes.SalesPeriodGroupMaster__c, spTgt);
                                    }
                                }
                                else {
                                    spTgt = {
                                        Id: "",
                                        Name: "",
                                        RoomTotal: Number(aggRes.RoomTotal)
                                    };
                                    salesPeriodMap.set(aggRes.SalesPeriodAndStage__c, spTgt);
                                }
                            }
                        }
                        catch (e_8_1) { e_8 = { error: e_8_1 }; }
                        finally {
                            try {
                                if (_f && !_f.done && (_j = _e.return)) _j.call(_e);
                            }
                            finally { if (e_8) throw e_8.error; }
                        }
                        query = "SELECT Id,Name,SalesPeriodGroupMaster__c,SalesPeriodGroupMaster__r.Name,RegisterPeriodStart__c,RegisterPeriodEnd__c,Lottery__c,ContractDate__c from SalesPeriodMaster__c WHERE Id in " + sfdc_service_1.SfdcService.getInSql(salesPeriodMasterId) + " ORDER BY RegisterPeriodStart__c,Lottery__c,ContractDate__c,SalesPeriodGroupMaster__c";
                        return [4, this.sfdcService.query(query).catch(function (err) {
                                throw err;
                            })];
                    case 2:
                        result = _m.sent();
                        optList = [];
                        salesPeriodGroupIdList = [];
                        try {
                            for (_g = tslib_1.__values(result.records), _h = _g.next(); !_h.done; _h = _g.next()) {
                                spm = _h.value;
                                if (spm.SalesPeriodGroupMaster__c) {
                                    if (salesPeriodGroupIdList.includes(spm.SalesPeriodGroupMaster__c)) {
                                        continue;
                                    }
                                    salesPeriodGroupIdList.push(spm.SalesPeriodGroupMaster__c);
                                    spTgt = salesPeriodMap.get(spm.SalesPeriodGroupMaster__c);
                                    if (!spTgt)
                                        continue;
                                    this.copySalesPeriodMaster(spTgt, spm);
                                    optList.push(spTgt);
                                }
                                else {
                                    spTgt = salesPeriodMap.get(spm.Id);
                                    if (!spTgt)
                                        continue;
                                    this.copySalesPeriodMaster(spTgt, spm);
                                    optList.push(spTgt);
                                }
                            }
                        }
                        catch (e_9_1) { e_9 = { error: e_9_1 }; }
                        finally {
                            try {
                                if (_h && !_h.done && (_k = _g.return)) _k.call(_g);
                            }
                            finally { if (e_9) throw e_9.error; }
                        }
                        rowIndex = 3;
                        try {
                            for (optList_1 = tslib_1.__values(optList), optList_1_1 = optList_1.next(); !optList_1_1.done; optList_1_1 = optList_1.next()) {
                                spInfo = optList_1_1.value;
                                rowIndex++;
                                this.worksheet.getCell("AC" + rowIndex).value = spInfo.Name;
                                this.worksheet.getCell("AG" + rowIndex).value = spInfo.RoomTotal ? spInfo.RoomTotal : "未定";
                                spInfo.RegisterPeriodStart__c = (_a = spInfo.RegisterPeriodStart__c) === null || _a === void 0 ? void 0 : _a.replace(/-/g, "/");
                                spInfo.RegisterPeriodEnd__c = (_b = spInfo.RegisterPeriodEnd__c) === null || _b === void 0 ? void 0 : _b.replace(/-/g, "/");
                                spInfo.Lottery__c = (_c = spInfo.Lottery__c) === null || _c === void 0 ? void 0 : _c.replace(/-/g, "/");
                                spInfo.ContractDate__c = (_d = spInfo.ContractDate__c) === null || _d === void 0 ? void 0 : _d.replace(/-/g, "/");
                                this.worksheet.getCell("AI" + rowIndex).value = "" + spInfo.RegisterPeriodStart__c + (spInfo.RegisterPeriodEnd__c ? "~" + spInfo.RegisterPeriodEnd__c : "");
                                this.worksheet.getCell("AN" + rowIndex).value = spInfo.Lottery__c;
                                this.worksheet.getCell("AP" + rowIndex).value = spInfo.ContractDate__c;
                            }
                        }
                        catch (e_10_1) { e_10 = { error: e_10_1 }; }
                        finally {
                            try {
                                if (optList_1_1 && !optList_1_1.done && (_l = optList_1.return)) _l.call(optList_1);
                            }
                            finally { if (e_10) throw e_10.error; }
                        }
                        return [2];
                }
            });
        });
    };
    WeeklyReportService.prototype.copySalesPeriodMaster = function (tgt, src) {
        tgt.Id = src.Id;
        tgt.Name = src.SalesPeriodGroupMaster__r ? src.SalesPeriodGroupMaster__r.Name : src.Name;
        tgt.Name = src.Name;
        tgt.SalesPeriodGroupMaster__c = src.SalesPeriodGroupMaster__c;
        tgt.SalesPeriodGroupMaster__r = src.SalesPeriodGroupMaster__r;
        tgt.RegisterPeriodStart__c = src.RegisterPeriodStart__c;
        tgt.RegisterPeriodEnd__c = src.RegisterPeriodEnd__c;
        tgt.Lottery__c = src.Lottery__c;
        tgt.ContractDate__c = src.ContractDate__c;
    };
    WeeklyReportService.prototype.sumBrillaClubInfo = function (opp, wenInfo, newsInfo) {
        var VisitDate = tools_service_1.Tools.strToDate(opp.VisitDate__c);
        if (!VisitDate)
            return;
        wenInfo.visitNum++;
        newsInfo.visitNum++;
        var ENewsletterAppDate = tools_service_1.Tools.strToDate(opp.ENewsletterAppDate__c);
        var NewsletterAppDate = tools_service_1.Tools.strToDate(opp.NewsletterAppDate__c);
        if (ENewsletterAppDate && ENewsletterAppDate < VisitDate && !opp.ENewsletterStopDate__c) {
            wenInfo.joinedNum++;
        }
        if (NewsletterAppDate && NewsletterAppDate < VisitDate && !opp.NewsletterStopDate__c) {
            newsInfo.joinedNum++;
        }
        if (opp.NewsletterAppRoute__c === "販売センター" && opp.JoinBCDate__c && opp.JoinBCDate__c === opp.VisitDate__c) {
            if (opp.ENewsletterProperty__c === opp.AccountId) {
                wenInfo.liveJoinNum++;
            }
            if (opp.NewsletterProperty__c === opp.AccountId) {
                newsInfo.liveJoinNum++;
            }
        }
    };
    WeeklyReportService.prototype.setBrilliaClub = function (WeeklyRptConfig) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var bcQuery, bcList, webInfo, newsLetterInfo, _a, _b, opp, rdtRes, BCRecordTypeId, bcRes, BCTotalInfo;
            var e_11, _c;
            return tslib_1.__generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        if (!this.theWeek || !this.worksheet)
                            return [2];
                        bcQuery = "SELECT Id,RecordTypeId,VisitDate__c,JoinBCDate__c,NewsletterAppRoute__c,ENewsletterProperty__c,NewsletterProperty__c,ENewsletterAppDate__c,ENewsletterStopDate__c,NewsletterAppDate__c,NewsletterStopDate__c,AccountId FROM Opportunity WHERE AccountId = '" + WeeklyRptConfig.projectId + "' AND RecordType.Name = 'BC' AND VisitDate__c >= " + this.theWeek.PeriodFrom__c + " AND VisitDate__c <= " + this.theWeek.PeriodTo__c;
                        return [4, this.sfdcService.query(bcQuery).catch(function (err) {
                                throw err;
                            })];
                    case 1:
                        bcList = _d.sent();
                        webInfo = new BrillaClubInfo();
                        newsLetterInfo = new BrillaClubInfo();
                        if (bcList.totalSize > 0) {
                            try {
                                for (_a = tslib_1.__values(bcList.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                    opp = _b.value;
                                    this.sumBrillaClubInfo(opp, webInfo, newsLetterInfo);
                                }
                            }
                            catch (e_11_1) { e_11 = { error: e_11_1 }; }
                            finally {
                                try {
                                    if (_b && !_b.done && (_c = _a.return)) _c.call(_a);
                                }
                                finally { if (e_11) throw e_11.error; }
                            }
                        }
                        this.worksheet.getCell("C49").value = webInfo.notJoinNum;
                        this.worksheet.getCell("E49").value = webInfo.liveJoinNum;
                        this.worksheet.getCell("G49").value = webInfo.notJoinNum === 0 ? 0 : webInfo.liveJoinNum / webInfo.notJoinNum;
                        this.worksheet.getCell("C50").value = newsLetterInfo.notJoinNum;
                        this.worksheet.getCell("E50").value = newsLetterInfo.liveJoinNum;
                        this.worksheet.getCell("G50").value =
                            newsLetterInfo.notJoinNum === 0 ? 0 : newsLetterInfo.liveJoinNum / newsLetterInfo.notJoinNum;
                        return [4, this.sfdcService.getRecordType("Opportunity", "BC")];
                    case 2:
                        rdtRes = _d.sent();
                        BCRecordTypeId = rdtRes.records[0].Id;
                        return [4, this.pgService
                                .query("\n    SELECT\n      t1.accountid as \"AccountId\",\n      array_to_string(array(SELECT count(1) FROM sfdc.opportunity where recordtypeid = t1.recordtypeid and accountid = t1.accountid and ENewsletterProperty__c = t1.accountid and VisitDate__c <= '" + this.theWeek.PeriodTo__c + "'),',') as \"webVisitTotal\",\n      array_to_string(array(SELECT count(1) FROM sfdc.opportunity where recordtypeid = t1.recordtypeid and accountid = t1.accountid and NewsletterProperty__c = t1.accountid and VisitDate__c <= '" + this.theWeek.PeriodTo__c + "'),',') as \"newsVisitTotal\",\n      array_to_string(array(SELECT count(1) FROM sfdc.opportunity where recordtypeid = t1.recordtypeid and accountid = t1.accountid and ENewsletterProperty__c = t1.accountid and JoinBCDate__c < VisitDate__c and VisitDate__c <= '" + this.theWeek.PeriodTo__c + "'),',') as \"webJoinedTotal\",\n      array_to_string(array(SELECT count(1) FROM sfdc.opportunity where recordtypeid = t1.recordtypeid and accountid = t1.accountid and NewsletterProperty__c = t1.accountid and JoinBCDate__c < VisitDate__c and VisitDate__c <= '" + this.theWeek.PeriodTo__c + "'),',') as \"newsJoinedTotal\",\n      array_to_string(array(SELECT count(1) FROM sfdc.opportunity where recordtypeid = t1.recordtypeid and accountid = t1.accountid and ENewsletterProperty__c = t1.accountid and JoinBCDate__c = VisitDate__c and VisitDate__c <= '" + this.theWeek.PeriodTo__c + "'),',') as \"webLiveJoinTotal\",\n      array_to_string(array(SELECT count(1) FROM sfdc.opportunity where recordtypeid = t1.recordtypeid and accountid = t1.accountid and NewsletterProperty__c = t1.accountid and JoinBCDate__c = VisitDate__c and VisitDate__c <= '" + this.theWeek.PeriodTo__c + "'),',') as \"newsLiveJoinTotal\"\n    FROM sfdc.opportunity t1 where t1.accountid = '" + WeeklyRptConfig.projectId + "' and t1.recordtypeid = '" + BCRecordTypeId + "' limit 1")
                                .catch(function (err) {
                                throw err;
                            })];
                    case 3:
                        bcRes = _d.sent();
                        BCTotalInfo = bcRes.rows[0];
                        if (!BCTotalInfo)
                            return [2];
                        this.worksheet.getCell("I49").value = Number(BCTotalInfo.webVisitTotal) - Number(BCTotalInfo.webJoinedTotal);
                        this.worksheet.getCell("K49").value = Number(BCTotalInfo.webLiveJoinTotal);
                        this.worksheet.getCell("M49").value =
                            Number(BCTotalInfo.webLiveJoinTotal) / (Number(BCTotalInfo.webVisitTotal) - Number(BCTotalInfo.webJoinedTotal));
                        this.worksheet.getCell("I50").value = Number(BCTotalInfo.newsVisitTotal) - Number(BCTotalInfo.newsJoinedTotal);
                        this.worksheet.getCell("K50").value = Number(BCTotalInfo.newsLiveJoinTotal);
                        this.worksheet.getCell("M50").value =
                            Number(BCTotalInfo.webLiveJoinTotal) / (Number(BCTotalInfo.newsVisitTotal) - Number(BCTotalInfo.newsJoinedTotal));
                        return [2];
                }
            });
        });
    };
    WeeklyReportService.prototype.initSebQuestionInfo = function () {
        this.theWeekSebQuestionInfo = new SebQuestionInfo();
        this.totalSebQuestionInfo = new SebQuestionInfo();
    };
    WeeklyReportService.prototype.sumSubQuestion = function (oppInfo) {
        if (!this.worksheet || !this.theWeek)
            return;
        if (!oppInfo.VisitDate__c)
            return;
        if (this.isTheWeek(oppInfo.VisitDate__c)) {
            this.setSubQuestion(oppInfo, this.theWeekSebQuestionInfo);
        }
        this.setSubQuestion(oppInfo, this.totalSebQuestionInfo);
    };
    WeeklyReportService.prototype.setSubQuestion = function (oppInfo, sebQuestionInfo) {
        var key = oppInfo.Opportunity__c + "_" + oppInfo.VisitDate__c;
        if (oppInfo.PriceImpression__c) {
            if (sebQuestionInfo.PriceImpression.includes(oppInfo.PriceImpression__c)) {
                sebQuestionInfo.weekPriceImpression.add(key);
            }
            sebQuestionInfo.weekPriceImpressionCount.add(key);
        }
        if (oppInfo.MREvaluation__c) {
            if (sebQuestionInfo.MREvaluation.includes(oppInfo.MREvaluation__c)) {
                sebQuestionInfo.weekMREvaluation.add(key);
            }
            sebQuestionInfo.weekMREvaluationCount.add(key);
        }
        if (oppInfo.LocalImpression__c && oppInfo.LocalImpression__c !== "まだ見ていない") {
            if (sebQuestionInfo.LocalImpression.includes(oppInfo.LocalImpression__c)) {
                sebQuestionInfo.weekLocalImpression.add(key);
            }
            sebQuestionInfo.weekLocalImpressionCount.add(key);
        }
        if (oppInfo.ExaminationDegree__c) {
            if (sebQuestionInfo.ExaminationDegree.includes(oppInfo.ExaminationDegree__c)) {
                sebQuestionInfo.weekExaminationDegree.add(key);
            }
            sebQuestionInfo.weekExaminationDegreeCount.add(key);
        }
    };
    WeeklyReportService.prototype.saveSubQuestion = function () {
        if (!this.worksheet || !this.theWeekSebQuestionInfo || !this.totalSebQuestionInfo)
            return;
        var sebQuestionInfo = this.theWeekSebQuestionInfo;
        this.worksheet.getCell("D53").value = sebQuestionInfo.weekPriceImpression.size;
        this.worksheet.getCell("E53").value =
            sebQuestionInfo.weekPriceImpressionCount.size === 0
                ? 0
                : sebQuestionInfo.weekPriceImpression.size / sebQuestionInfo.weekPriceImpressionCount.size;
        this.worksheet.getCell("K53").value = sebQuestionInfo.weekLocalImpression.size;
        this.worksheet.getCell("L53").value =
            sebQuestionInfo.weekLocalImpressionCount.size === 0
                ? 0
                : sebQuestionInfo.weekLocalImpression.size / sebQuestionInfo.weekLocalImpressionCount.size;
        this.worksheet.getCell("D54").value = sebQuestionInfo.weekExaminationDegree.size;
        this.worksheet.getCell("E54").value =
            sebQuestionInfo.weekExaminationDegreeCount.size === 0
                ? 0
                : sebQuestionInfo.weekExaminationDegree.size / sebQuestionInfo.weekExaminationDegreeCount.size;
        this.worksheet.getCell("K54").value = sebQuestionInfo.weekMREvaluation.size;
        this.worksheet.getCell("L54").value =
            sebQuestionInfo.weekMREvaluationCount.size === 0
                ? 0
                : sebQuestionInfo.weekMREvaluation.size / sebQuestionInfo.weekMREvaluationCount.size;
        sebQuestionInfo = this.totalSebQuestionInfo;
        this.worksheet.getCell("F53").value = sebQuestionInfo.weekPriceImpression.size;
        this.worksheet.getCell("G53").value =
            sebQuestionInfo.weekPriceImpressionCount.size === 0
                ? 0
                : sebQuestionInfo.weekPriceImpression.size / sebQuestionInfo.weekPriceImpressionCount.size;
        this.worksheet.getCell("M53").value = sebQuestionInfo.weekLocalImpression.size;
        this.worksheet.getCell("N53").value =
            sebQuestionInfo.weekLocalImpressionCount.size === 0
                ? 0
                : sebQuestionInfo.weekLocalImpression.size / sebQuestionInfo.weekLocalImpressionCount.size;
        this.worksheet.getCell("F54").value = sebQuestionInfo.weekExaminationDegree.size;
        this.worksheet.getCell("G54").value =
            sebQuestionInfo.weekExaminationDegreeCount.size === 0
                ? 0
                : sebQuestionInfo.weekExaminationDegree.size / sebQuestionInfo.weekExaminationDegreeCount.size;
        this.worksheet.getCell("M54").value = sebQuestionInfo.weekMREvaluation.size;
        this.worksheet.getCell("N54").value =
            sebQuestionInfo.weekMREvaluationCount.size === 0
                ? 0
                : sebQuestionInfo.weekMREvaluation.size / sebQuestionInfo.weekMREvaluationCount.size;
    };
    WeeklyReportService.prototype.setOOSInfo = function (weeklyRptCfg) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var query, result, oosInfo, _a, _b, opp, ResponseDate, VisitDate, _c, _d, contract, ContractDate;
            var e_12, _e, e_13, _f;
            return tslib_1.__generator(this, function (_g) {
                switch (_g.label) {
                    case 0:
                        if (!this.worksheet || !this.theWeek)
                            return [2];
                        query = "SELECT Id,StageName,ResponseDate__c,VisitDate__c,(SELECT Id,ContractDate__c,CancellationDate__c FROM Contracts__r) FROM Opportunity WHERE Accountid = '" + weeklyRptCfg.projectId + "' AND StageName = 'OSS\uFF08\u5951\u7D04\u524D\uFF09' and ResponseDate__c != null";
                        return [4, this.sfdcService.query(query).catch(function (err) {
                                throw err;
                            })];
                    case 1:
                        result = _g.sent();
                        if (result.totalSize === 0)
                            return [2];
                        oosInfo = {
                            beforeNum: 0,
                            visitNum: 0,
                            get beforePercent() {
                                if (this.beforeNum === 0) {
                                    return 0;
                                }
                                else {
                                    return this.visitNum / this.beforeNum;
                                }
                            },
                            beforeContractNum: 0,
                            contractNum: 0,
                            get beforeContractPercent() {
                                if (this.beforeContractNum === 0) {
                                    return 0;
                                }
                                else {
                                    return this.contractNum / this.beforeContractNum;
                                }
                            }
                        };
                        try {
                            for (_a = tslib_1.__values(result.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                opp = _b.value;
                                ResponseDate = tools_service_1.Tools.strToDate(opp.ResponseDate__c);
                                if (!ResponseDate)
                                    continue;
                                VisitDate = tools_service_1.Tools.strToDate(opp.VisitDate__c);
                                if (!VisitDate || ResponseDate < VisitDate) {
                                    oosInfo.beforeNum++;
                                }
                                if (VisitDate && ResponseDate < VisitDate) {
                                    oosInfo.visitNum++;
                                }
                                if (!opp.Contracts__r)
                                    continue;
                                try {
                                    for (_c = (e_13 = void 0, tslib_1.__values(opp.Contracts__r.records)), _d = _c.next(); !_d.done; _d = _c.next()) {
                                        contract = _d.value;
                                        ContractDate = tools_service_1.Tools.strToDate(contract.ContractDate__c);
                                        if (!ContractDate || ResponseDate < ContractDate) {
                                            oosInfo.beforeContractNum++;
                                        }
                                        if (ContractDate && ResponseDate < ContractDate) {
                                            oosInfo.contractNum++;
                                        }
                                    }
                                }
                                catch (e_13_1) { e_13 = { error: e_13_1 }; }
                                finally {
                                    try {
                                        if (_d && !_d.done && (_f = _c.return)) _f.call(_c);
                                    }
                                    finally { if (e_13) throw e_13.error; }
                                }
                            }
                        }
                        catch (e_12_1) { e_12 = { error: e_12_1 }; }
                        finally {
                            try {
                                if (_b && !_b.done && (_e = _a.return)) _e.call(_a);
                            }
                            finally { if (e_12) throw e_12.error; }
                        }
                        this.worksheet.getCell("C57").value = oosInfo.beforeNum;
                        this.worksheet.getCell("D57").value = oosInfo.visitNum;
                        this.worksheet.getCell("F57").value = oosInfo.beforePercent;
                        this.worksheet.getCell("I57").value = oosInfo.beforeContractNum;
                        this.worksheet.getCell("D57").value = oosInfo.contractNum;
                        this.worksheet.getCell("D57").value = oosInfo.beforeContractPercent;
                        return [2];
                }
            });
        });
    };
    WeeklyReportService.prototype.sumTableInfo = function (record, weekDateMap) {
        if (record.Id === record.Opportunity__r.OpportunityHistory__c) {
            this.setWeekTableInfo(record.Opportunity__r, weekDateMap, "InquiryDate__c", "InquiryDate__c", "Id");
            this.setWeekTableInfo(record.Opportunity__r, weekDateMap, "FirstReturnVisitDate__c", "ReturnVisitDate__c", "Id");
        }
        this.setWeekTableInfo(record, weekDateMap, "OSSRelationLatestDate__c", undefined, "Opportunity__c");
        this.setWeekTableInfo(record, weekDateMap, "VisitDate__c", undefined, "Opportunity__c");
        this.setWeekTableInfo(record, weekDateMap, "ReturnLatestVisitDate__c", undefined, "Opportunity__c");
        this.setWeekTableInfo(record, weekDateMap, "RegisterDate__c", undefined, "Opportunity__c");
        this.setWeekTableInfo(record, weekDateMap, "RegisterCancelDate__c", undefined, "Opportunity__c");
        this.setWeekTableInfo(record, weekDateMap, "ReserveDate__c", "RegisterCancelDate__c", "Opportunity__c");
    };
    WeeklyReportService.prototype.saveTableInfo = function (weeklyRptConfig, weekDateMap, queryStartDate, queryEndDate) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var contractQuery, contractList, _a, _b, contract, conRes, conTotalRes, salesMstQuery, salesRes, _c, _d, salesCountInfo, weekDate_1, footstayRes, _e, _f, weekDate_2, weekDate, visitCount, theWeekDate, InquiryVisitNum, footCtr;
            var e_14, _g, e_15, _h, e_16, _j;
            return tslib_1.__generator(this, function (_k) {
                switch (_k.label) {
                    case 0:
                        if (!this.worksheet)
                            return [2];
                        contractQuery = "SELECT Id,ContractDate__c,ApplicationDate__c,CancelDate__c,CancellationDate__c,DeliverDate__c,ReserveDate__c,ResponseKinds__c FROM Contract WHERE AccountId = '" + weeklyRptConfig.projectId + "' AND ((ContractDate__c >= " + queryStartDate + " AND ContractDate__c <= " + queryEndDate + ") OR (ApplicationDate__c >= " + queryStartDate + " AND ApplicationDate__c <= " + queryEndDate + ") OR (CancellationDate__c >= " + queryStartDate + " AND CancellationDate__c <= " + queryEndDate + ") OR (DeliverDate__c >= " + queryStartDate + " AND DeliverDate__c <= " + queryEndDate + ") OR (ReserveDate__c >= " + queryStartDate + " AND ReserveDate__c <= " + queryEndDate + ")  OR (CancelDate__c >= " + queryStartDate + " AND CancelDate__c <= " + queryEndDate + "))";
                        return [4, this.sfdcService.query(contractQuery).catch(function (err) {
                                throw err;
                            })];
                    case 1:
                        contractList = _k.sent();
                        if (contractList.totalSize > 0) {
                            try {
                                for (_a = tslib_1.__values(contractList.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                    contract = _b.value;
                                    if (contract.ResponseKinds__c === "予約") {
                                        this.setWeekTableInfo(contract, weekDateMap, "ReserveDate__c");
                                        this.setWeekTableInfo(contract, weekDateMap, "CancelDate__c", "ReserveCancell");
                                    }
                                    else if (contract.ResponseKinds__c === "申込") {
                                        this.setWeekTableInfo(contract, weekDateMap, "ApplicationDate__c");
                                        this.setWeekTableInfo(contract, weekDateMap, "ApplicationDate__c", "ReserveCancell");
                                        this.setWeekTableInfo(contract, weekDateMap, "CancelDate__c", "ApplicationCancell");
                                    }
                                    else if (contract.ResponseKinds__c === "契約") {
                                        this.setWeekTableInfo(contract, weekDateMap, "ContractDate__c");
                                        this.setWeekTableInfo(contract, weekDateMap, "CancellationDate__c", "ContractCancell");
                                        this.setWeekTableInfo(contract, weekDateMap, "ContractDate__c", "ReserveCancell");
                                        this.setWeekTableInfo(contract, weekDateMap, "ContractDate__c", "ApplicationCancell");
                                    }
                                    else if (contract.ResponseKinds__c === "引渡") {
                                        this.setWeekTableInfo(contract, weekDateMap, "DeliverDate__c");
                                    }
                                }
                            }
                            catch (e_14_1) { e_14 = { error: e_14_1 }; }
                            finally {
                                try {
                                    if (_b && !_b.done && (_g = _a.return)) _g.call(_a);
                                }
                                finally { if (e_14) throw e_14.error; }
                            }
                        }
                        return [4, this.pgService
                                .query("\n      SELECT\n        t1.accountid as \"AccountId\",\n        array_to_string(array(SELECT count(1) FROM sfdc.Contract where accountid = t1.accountid and ReserveDate__c <= '" + queryEndDate + "' and ResponseKinds__c = '\u4E88\u7D04'),',') as \"ReserveDate__c\",\n        array_to_string(array(SELECT count(1) FROM sfdc.Contract where accountid = t1.accountid and CancellationDate__c <= '" + queryEndDate + "' and ResponseKinds__c = '\u4E88\u7D04'),',') as \"ReserveCancell\",\n        array_to_string(array(SELECT count(1) FROM sfdc.Contract where accountid = t1.accountid and ApplicationDate__c <= '" + queryEndDate + "' and ResponseKinds__c = '\u7533\u8FBC'),',') as \"ApplicationDate__c\",\n        array_to_string(array(SELECT count(1) FROM sfdc.Contract where accountid = t1.accountid and CancellationDate__c <= '" + queryEndDate + "' and ResponseKinds__c = '\u7533\u8FBC'),',') as \"ApplicationCancell\",\n        array_to_string(array(SELECT count(1) FROM sfdc.Contract where accountid = t1.accountid and ContractDate__c <= '" + queryEndDate + "' and ResponseKinds__c = '\u5951\u7D04'),',') as \"ContractDate__c\",\n        array_to_string(array(SELECT count(1) FROM sfdc.Contract where accountid = t1.accountid and CancellationDate__c <= '" + queryEndDate + "' and ResponseKinds__c = '\u5951\u7D04'),',') as \"ContractCancell\",\n        array_to_string(array(SELECT count(1) FROM sfdc.Contract where accountid = t1.accountid and DeliverDate__c <= '" + queryEndDate + "' and ResponseKinds__c = '\u5F15\u6E21'),',') as \"DeliverDate__c\"\n      FROM sfdc.Contract t1 where t1.AccountId = '" + weeklyRptConfig.projectId + "' limit 1")
                                .catch(function (err) {
                                throw err;
                            })];
                    case 2:
                        conRes = _k.sent();
                        conTotalRes = conRes.rows[0];
                        salesMstQuery = "select SalesPeriodAndStage__r.Lottery__c,count(id) RoomCount from Room__c where SalesPeriodAndStage__c != null and Project__c = '" + weeklyRptConfig.projectId + "' and (SalesPeriodAndStage__r.Lottery__c >= " + queryStartDate + " and SalesPeriodAndStage__r.Lottery__c <= " + queryEndDate + ") group by SalesPeriodAndStage__r.Lottery__c";
                        return [4, this.sfdcService.query(salesMstQuery).catch(function (err) {
                                throw err;
                            })];
                    case 3:
                        salesRes = _k.sent();
                        if (salesRes.totalSize > 0) {
                            try {
                                for (_c = tslib_1.__values(salesRes.records), _d = _c.next(); !_d.done; _d = _c.next()) {
                                    salesCountInfo = _d.value;
                                    weekDate_1 = weekDateMap.get(salesCountInfo.Lottery__c);
                                    if (weekDate_1)
                                        weekDate_1.NoneNum += salesCountInfo.RoomCount;
                                    weekDate_1 = this.getWeekByDate(weekDateMap, salesCountInfo.Lottery__c);
                                    if (weekDate_1)
                                        weekDate_1.NoneNum += salesCountInfo.RoomCount;
                                }
                            }
                            catch (e_15_1) { e_15 = { error: e_15_1 }; }
                            finally {
                                try {
                                    if (_d && !_d.done && (_h = _c.return)) _h.call(_c);
                                }
                                finally { if (e_15) throw e_15.error; }
                            }
                        }
                        if (!this.footstay) return [3, 5];
                        return [4, this.pgService
                                .query("\n      SELECT\n        t1.accountid as \"AccountId\",\n        array_to_string(array(SELECT count(1) FROM sfdc.Contract where accountid = t1.accountid and ContractDate__c <= '" + queryEndDate + "' and ContractDate__c >= '" + date_and_time_1.default.format(this.footstay.month3Start, "YYYY-MM-DD") + "' and (CancellationDate__c is null or CancellationDate__c > '" + queryEndDate + "')),',') as \"month3Contract\",\n      array_to_string(array(SELECT count(1) FROM sfdc.Contract where accountid = t1.accountid and ContractDate__c <= '" + queryEndDate + "' and ContractDate__c >= '" + date_and_time_1.default.format(this.footstay.yearStart, "YYYY-MM-DD") + "' and (CancellationDate__c is null or CancellationDate__c > '" + queryEndDate + "')),',') as \"yearContract\"\n      FROM sfdc.Contract t1 where t1.AccountId = '" + weeklyRptConfig.projectId + "' limit 1")
                                .catch(function (err) {
                                throw err;
                            })];
                    case 4:
                        footstayRes = _k.sent();
                        _k.label = 5;
                    case 5:
                        try {
                            for (_e = tslib_1.__values(weekDateMap.values()), _f = _e.next(); !_f.done; _f = _e.next()) {
                                weekDate_2 = _f.value;
                                this.worksheet.getCell(34, weekDate_2.xindex).value = weekDate_2.InquiryDate__c.size;
                                this.worksheet.getCell(35, weekDate_2.xindex).value = weekDate_2.OSSRelationLatestDate__c.size;
                                this.worksheet.getCell(36, weekDate_2.xindex).value = weekDate_2.VisitDate__c.size;
                                this.worksheet.getCell(37, weekDate_2.xindex).value = weekDate_2.ReturnVisitDate__c.size;
                                this.worksheet.getCell(38, weekDate_2.xindex).value = weekDate_2.ReturnLatestVisitDate__c.size;
                                if (weekDate_2.date) {
                                    this.worksheet.getCell(39, weekDate_2.xindex).value = weekDate_2.RegisterDate__c.size;
                                    this.worksheet.getCell(39, weekDate_2.xindex + 1).value = -weekDate_2.RegisterCancelDate__c.size;
                                    this.worksheet.getCell(40, weekDate_2.xindex).value = weekDate_2.ReserveDate__c;
                                    this.worksheet.getCell(40, weekDate_2.xindex + 1).value = -weekDate_2.ReserveCancell;
                                    this.worksheet.getCell(41, weekDate_2.xindex).value = weekDate_2.ApplicationDate__c;
                                    this.worksheet.getCell(41, weekDate_2.xindex + 1).value = -weekDate_2.ApplicationCancell;
                                    this.worksheet.getCell(42, weekDate_2.xindex).value = weekDate_2.ContractDate__c;
                                    this.worksheet.getCell(42, weekDate_2.xindex + 1).value = -weekDate_2.ContractCancell;
                                    this.worksheet.getCell(43, weekDate_2.xindex).value = weekDate_2.NoneNum;
                                    this.worksheet.getCell(43, weekDate_2.xindex + 1).value = -weekDate_2.ReserveDate__c - weekDate_2.ApplicationDate__c;
                                }
                                else {
                                    this.worksheet.getCell(39, weekDate_2.xindex).value =
                                        weekDate_2.RegisterDate__c.size - weekDate_2.RegisterCancelDate__c.size;
                                    this.worksheet.getCell(40, weekDate_2.xindex).value = weekDate_2.ReserveDate__c - weekDate_2.ReserveCancell;
                                    this.worksheet.getCell(41, weekDate_2.xindex).value = weekDate_2.ApplicationDate__c - weekDate_2.ApplicationCancell;
                                    this.worksheet.getCell(42, weekDate_2.xindex).value = weekDate_2.ContractDate__c - weekDate_2.ContractCancell;
                                    this.worksheet.getCell(43, weekDate_2.xindex).value =
                                        weekDate_2.NoneNum - weekDate_2.ReserveDate__c - weekDate_2.ApplicationDate__c;
                                }
                                this.worksheet.getCell(44, weekDate_2.xindex).value = weekDate_2.DeliverDate__c;
                            }
                        }
                        catch (e_16_1) { e_16 = { error: e_16_1 }; }
                        finally {
                            try {
                                if (_f && !_f.done && (_j = _e.return)) _j.call(_e);
                            }
                            finally { if (e_16) throw e_16.error; }
                        }
                        if (conTotalRes) {
                            this.worksheet.getCell("AI40").value = conTotalRes.ReserveDate__c - conTotalRes.ReserveCancell;
                            this.worksheet.getCell("AI41").value = conTotalRes.ApplicationDate__c - conTotalRes.ApplicationCancell;
                            this.worksheet.getCell("AI42").value = conTotalRes.ContractDate__c - conTotalRes.ContractCancell;
                            this.worksheet.getCell("AI44").value = conTotalRes.DeliverDate__c;
                            this.worksheet.getCell("S40").value =
                                Number(this.worksheet.getCell("AI40").value) - Number(this.worksheet.getCell("AH40").value);
                            this.worksheet.getCell("S41").value =
                                Number(this.worksheet.getCell("AI41").value) - Number(this.worksheet.getCell("AH41").value);
                            this.worksheet.getCell("S42").value =
                                Number(this.worksheet.getCell("AI42").value) - Number(this.worksheet.getCell("AH42").value);
                            this.worksheet.getCell("S44").value = conTotalRes.DeliverDate__c - Number(this.worksheet.getCell("AH44").value);
                        }
                        weekDate = weekDateMap.get("累計");
                        if (weekDate) {
                            this.worksheet.getCell("S34").value = weekDate.InquiryDate__c.size - Number(this.worksheet.getCell("AH34").value);
                            this.worksheet.getCell("S35").value =
                                weekDate.OSSRelationLatestDate__c.size - Number(this.worksheet.getCell("AH35").value);
                            this.worksheet.getCell("S36").value = weekDate.VisitDate__c.size - Number(this.worksheet.getCell("AH36").value);
                            this.worksheet.getCell("S37").value =
                                weekDate.ReturnVisitDate__c.size - Number(this.worksheet.getCell("AH37").value);
                            this.worksheet.getCell("S38").value =
                                weekDate.ReturnLatestVisitDate__c.size - Number(this.worksheet.getCell("AH38").value);
                            this.worksheet.getCell("S39").value = weekDate.RegisterDate__c.size - weekDate.RegisterCancelDate__c.size;
                            Number(this.worksheet.getCell("AH39").value);
                            visitCount = weekDate.InquiryDate__c.size + weekDate.ReturnVisitDate__c.size + weekDate.ReturnLatestVisitDate__c.size;
                            this.worksheet.getCell("AL33").value = weekDate.InquiryDate__c.size / weeklyRptConfig.selectWeek.weekNum;
                            this.worksheet.getCell("AL34").value = visitCount / weeklyRptConfig.selectWeek.weekNum;
                            this.worksheet.getCell("AL35").value = weekDate.ReturnVisitDate__c.size / visitCount;
                            this.worksheet.getCell("AL36").value = weekDate.ReturnLatestVisitDate__c.size / visitCount;
                            this.worksheet.getCell("AL37").value = weekDate.InquiryVisitNum.size;
                            this.worksheet.getCell("AL38").value = weekDate.InquiryVisitNum.size / visitCount;
                            this.worksheet.getCell("AL39").value = weekDate.InquiryVisitNum.size / weekDate.VisitDate__c.size;
                            theWeekDate = weekDateMap.get("今週");
                            InquiryVisitNum = theWeekDate && theWeekDate.InquiryVisitNum ? theWeekDate.InquiryVisitNum.size : 0;
                            this.worksheet.getCell("AL40").value = InquiryVisitNum;
                            this.worksheet.getCell("AL41").value = InquiryVisitNum / weekDate.VisitDate__c.size;
                        }
                        if (footstayRes && this.footstay) {
                            footCtr = footstayRes.rows[0];
                            this.worksheet.getCell("AL42").value =
                                Number(footCtr.month3Contract) / (this.footstay.month3Visit.size === 0 ? 1 : this.footstay.month3Visit.size);
                            this.worksheet.getCell("AL43").value =
                                Number(footCtr.yearContract) / (this.footstay.yearVisit.size === 0 ? 1 : this.footstay.yearVisit.size);
                            this.worksheet.getCell("AL44").value =
                                Number(this.worksheet.getCell("AI42").value) / Number(this.worksheet.getCell("AI36").value);
                        }
                        return [2];
                }
            });
        });
    };
    WeeklyReportService.prototype.setWeekTableInfo = function (sfdcObj, weekDateMap, filedName, weekDateFiledName, keyFiled) {
        if (!sfdcObj[filedName])
            return;
        var dateKey = sfdcObj[filedName].substring(0, 10);
        var weekTable = weekDateMap.get(dateKey);
        if (weekTable) {
            this.setWeekTableVal(weekTable, sfdcObj, keyFiled, weekDateFiledName, filedName, dateKey);
        }
        weekTable = this.getWeekByDate(weekDateMap, dateKey);
        if (weekTable) {
            this.setWeekTableVal(weekTable, sfdcObj, keyFiled, weekDateFiledName, filedName, dateKey);
        }
        weekTable = weekDateMap.get("累計");
        if (weekTable && keyFiled) {
            this.setWeekTableVal(weekTable, sfdcObj, keyFiled, weekDateFiledName, filedName, dateKey);
        }
    };
    WeeklyReportService.prototype.isTheWeek = function (dateKey) {
        var compareDate = tools_service_1.Tools.strToDate(dateKey);
        if (this.theWeek &&
            this.theWeek.PeriodFrom &&
            this.theWeek.PeriodTo &&
            compareDate &&
            compareDate >= this.theWeek.PeriodFrom &&
            compareDate <= this.theWeek.PeriodTo) {
            return true;
        }
        else {
            return false;
        }
    };
    WeeklyReportService.prototype.getWeekByDate = function (weekDateMap, dateKey) {
        var weekTable;
        var compareDate = tools_service_1.Tools.strToDate(dateKey);
        if (this.beforeLastWeek &&
            this.beforeLastWeek.PeriodFrom &&
            this.beforeLastWeek.PeriodTo &&
            compareDate &&
            compareDate >= this.beforeLastWeek.PeriodFrom &&
            compareDate <= this.beforeLastWeek.PeriodTo) {
            weekTable = weekDateMap.get("先々週");
        }
        else if (this.lastWeek &&
            this.lastWeek.PeriodFrom &&
            this.lastWeek.PeriodTo &&
            compareDate &&
            compareDate >= this.lastWeek.PeriodFrom &&
            compareDate <= this.lastWeek.PeriodTo) {
            weekTable = weekDateMap.get("先週");
        }
        else if (this.theWeek &&
            this.theWeek.PeriodFrom &&
            this.theWeek.PeriodTo &&
            compareDate &&
            compareDate >= this.theWeek.PeriodFrom &&
            compareDate <= this.theWeek.PeriodTo) {
            weekTable = weekDateMap.get("今週");
        }
        return weekTable;
    };
    WeeklyReportService.prototype.setWeekTableVal = function (weekTable, sfdcObj, keyFiled, weekDateFiledName, filedName, dateKey) {
        if (keyFiled) {
            var setKey = sfdcObj[keyFiled] + "_" + dateKey;
            weekTable[weekDateFiledName ? weekDateFiledName : filedName].add(setKey);
        }
        else {
            weekTable[weekDateFiledName ? weekDateFiledName : filedName]++;
        }
        if (weekTable.weekName === "今週" || weekTable.weekName === "累計") {
            if (keyFiled && filedName === "VisitDate__c" && sfdcObj.InquiryDate__c) {
                var setKey = sfdcObj[keyFiled] + "_" + dateKey + "_" + sfdcObj.VisitDate__c;
                weekTable.InquiryVisitNum.add(setKey);
            }
        }
    };
    WeeklyReportService.prototype.saveContractGoal = function (weeklyRptConfig) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var startDate, endDate, monthlyOptMap, index, monthMst, theDate, dateFormat, contractSql, resContract, _a, _b, row, opt, cancelSql, resCancel, _c, _d, row, opt, totalSql, resTotal, total, _e, _f, monthly, _g, _h, monthly;
            var e_17, _j, e_18, _k, e_19, _l, e_20, _m;
            return tslib_1.__generator(this, function (_o) {
                switch (_o.label) {
                    case 0:
                        if (!this.worksheet)
                            return [2];
                        startDate = tools_service_1.Tools.strToDate(weeklyRptConfig.monthMasterList[0].Month__c + "/" + "01");
                        endDate = tools_service_1.Tools.strToDate(weeklyRptConfig.monthMasterList[weeklyRptConfig.monthMasterList.length - 1].Month__c + "/" + "01");
                        if (!startDate || !endDate)
                            return [2];
                        endDate = date_and_time_1.default.addMonths(endDate, 1);
                        monthlyOptMap = new Map();
                        for (index = 0; index < weeklyRptConfig.monthMasterList.length; index++) {
                            monthMst = weeklyRptConfig.monthMasterList[index];
                            theDate = tools_service_1.Tools.strToDate(monthMst.Month__c + "/" + "01");
                            if (!theDate)
                                continue;
                            dateFormat = date_and_time_1.default.format(theDate, "YYYY/MM");
                            monthlyOptMap.set(dateFormat, {
                                month: dateFormat,
                                rowIndex: 49 + index,
                                monthMst: monthMst,
                                monthlyContract: null,
                                monthlyCancel: null,
                                contractTotal: null
                            });
                        }
                        contractSql = "select to_char(contractdate__c,'YYYY/MM') monthly,count(1) countnum from sfdc.contract where (contractdate__c >= '" + date_and_time_1.default.format(startDate, "YYYY-MM-DD") + "' and contractdate__c < '" + date_and_time_1.default.format(endDate, "YYYY-MM-DD") + "') group by to_char(contractdate__c,'YYYY/MM')";
                        return [4, this.pgService.query(contractSql).catch(function (err) {
                                throw err;
                            })];
                    case 1:
                        resContract = _o.sent();
                        try {
                            for (_a = tslib_1.__values(resContract.rows), _b = _a.next(); !_b.done; _b = _a.next()) {
                                row = _b.value;
                                opt = monthlyOptMap.get(row.monthly);
                                if (!opt)
                                    continue;
                                opt.monthlyContract = Number(row.countnum);
                            }
                        }
                        catch (e_17_1) { e_17 = { error: e_17_1 }; }
                        finally {
                            try {
                                if (_b && !_b.done && (_j = _a.return)) _j.call(_a);
                            }
                            finally { if (e_17) throw e_17.error; }
                        }
                        cancelSql = "select to_char(cancellationDate__c,'YYYY/MM') monthly,count(1) countnum from sfdc.contract where (cancellationDate__c >= '" + date_and_time_1.default.format(startDate, "YYYY-MM-DD") + "' and cancellationDate__c < '" + date_and_time_1.default.format(endDate, "YYYY-MM-DD") + "') group by to_char(cancellationDate__c,'YYYY/MM')";
                        return [4, this.pgService.query(cancelSql).catch(function (err) {
                                throw err;
                            })];
                    case 2:
                        resCancel = _o.sent();
                        try {
                            for (_c = tslib_1.__values(resCancel.rows), _d = _c.next(); !_d.done; _d = _c.next()) {
                                row = _d.value;
                                opt = monthlyOptMap.get(row.monthly);
                                if (!opt)
                                    continue;
                                opt.monthlyCancel = -Number(row.countnum);
                            }
                        }
                        catch (e_18_1) { e_18 = { error: e_18_1 }; }
                        finally {
                            try {
                                if (_d && !_d.done && (_k = _c.return)) _k.call(_c);
                            }
                            finally { if (e_18) throw e_18.error; }
                        }
                        totalSql = "select count(1) total from sfdc.contract\n    where contractdate__c < '" + date_and_time_1.default.format(startDate, "YYYY-MM-DD") + "' and (cancellationDate__c is null or cancellationDate__c >= '" + date_and_time_1.default.format(startDate, "YYYY-MM-DD") + "')";
                        return [4, this.pgService.query(totalSql).catch(function (err) {
                                throw err;
                            })];
                    case 3:
                        resTotal = _o.sent();
                        total = Number(resTotal.rows[0].total);
                        try {
                            for (_e = tslib_1.__values(monthlyOptMap.values()), _f = _e.next(); !_f.done; _f = _e.next()) {
                                monthly = _f.value;
                                if (!monthly.monthlyContract && !monthly.monthlyCancel)
                                    continue;
                                total =
                                    total +
                                        (monthly.monthlyContract ? monthly.monthlyContract : 0) +
                                        (monthly.monthlyCancel ? monthly.monthlyCancel : 0);
                                monthly.contractTotal = total;
                            }
                        }
                        catch (e_19_1) { e_19 = { error: e_19_1 }; }
                        finally {
                            try {
                                if (_f && !_f.done && (_l = _e.return)) _l.call(_e);
                            }
                            finally { if (e_19) throw e_19.error; }
                        }
                        try {
                            for (_g = tslib_1.__values(monthlyOptMap.values()), _h = _g.next(); !_h.done; _h = _g.next()) {
                                monthly = _h.value;
                                this.worksheet.getCell(monthly.rowIndex, 37).value = monthly.month;
                                this.worksheet.getCell(monthly.rowIndex, 39).value = monthly.monthlyContract;
                                this.worksheet.getCell(monthly.rowIndex, 40).value = monthly.monthlyCancel;
                                this.worksheet.getCell(monthly.rowIndex, 41).value = monthly.contractTotal;
                                this.worksheet.getCell(monthly.rowIndex, 43).value = monthly.monthMst.GoalContractMonthWeeklyReport__c;
                                this.worksheet.getCell(monthly.rowIndex, 45).value = monthly.monthMst.GoalContractTotalWeeklyReport__c;
                                this.worksheet.getCell(monthly.rowIndex, 49).value = monthly.monthMst.GoalCustomerMonthWeeklyReport__c;
                            }
                        }
                        catch (e_20_1) { e_20 = { error: e_20_1 }; }
                        finally {
                            try {
                                if (_h && !_h.done && (_m = _g.return)) _m.call(_g);
                            }
                            finally { if (e_20) throw e_20.error; }
                        }
                        return [2];
                }
            });
        });
    };
    WeeklyReportService.prototype.initMonthlyCustomerTotalGoal = function (weeklyRptConfig) {
        this.monthlyCustomerTotalGoal = new Map();
        for (var index = 0; index < weeklyRptConfig.monthMasterList.length; index++) {
            var monthMst = weeklyRptConfig.monthMasterList[index];
            var theDate = tools_service_1.Tools.strToDate(monthMst.Month__c + "/" + "01");
            if (!theDate)
                continue;
            var dateFormat = date_and_time_1.default.format(theDate, "YYYY/MM");
            this.monthlyCustomerTotalGoal.set(dateFormat, {
                month: dateFormat,
                rowIndex: 49 + index,
                visitCount: new Set()
            });
        }
    };
    WeeklyReportService.prototype.setMonthlyCustomerTotalGoal = function (record) {
        var e_21, _a;
        var countFiled = ["VisitDate__c"];
        try {
            for (var countFiled_1 = tslib_1.__values(countFiled), countFiled_1_1 = countFiled_1.next(); !countFiled_1_1.done; countFiled_1_1 = countFiled_1.next()) {
                var filed = countFiled_1_1.value;
                var theDate = tools_service_1.Tools.strToDate(record[filed]);
                if (!theDate)
                    continue;
                var dateFormat = date_and_time_1.default.format(theDate, "YYYY/MM");
                var monthlyCustomer = this.monthlyCustomerTotalGoal.get(dateFormat);
                if (monthlyCustomer) {
                    monthlyCustomer.visitCount.add(record.Opportunity__c + "_" + record[filed]);
                }
            }
        }
        catch (e_21_1) { e_21 = { error: e_21_1 }; }
        finally {
            try {
                if (countFiled_1_1 && !countFiled_1_1.done && (_a = countFiled_1.return)) _a.call(countFiled_1);
            }
            finally { if (e_21) throw e_21.error; }
        }
    };
    WeeklyReportService.prototype.saveMonthlyCustomerTotalGoal = function () {
        var e_22, _a;
        if (!this.worksheet)
            return;
        try {
            for (var _b = tslib_1.__values(this.monthlyCustomerTotalGoal.values()), _c = _b.next(); !_c.done; _c = _b.next()) {
                var monthly = _c.value;
                this.worksheet.getCell(monthly.rowIndex, 47).value = monthly.visitCount.size;
            }
        }
        catch (e_22_1) { e_22 = { error: e_22_1 }; }
        finally {
            try {
                if (_c && !_c.done && (_a = _b.return)) _a.call(_b);
            }
            finally { if (e_22) throw e_22.error; }
        }
    };
    WeeklyReportService.prototype.setSupplystatus = function (weeklyRptConfig) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var query, totalInfo, nowInfo, supplyedInfo, supplyedRes, lastSalesPeriodAndStageId, _a, _b, room, optList, optList_2, optList_2_1, optInfo, res, countData;
            var e_23, _c, e_24, _d;
            return tslib_1.__generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        if (!this.worksheet)
                            return [2];
                        query = "SELECT SalesPeriodAndStage__c,SalesPeriodAndStage__r.Lottery__c,SalesPeriodAndStage__r.SalesPeriodGroupMaster__c,Project__c,(select id,ResponseKinds__c,ReserveDate__c,ContractDate__c,CancellationDate__c from Contacts__r order by ContractNumber DESC limit 1),ForAndnonForSaleStore__c FROM Room__c where Project__c = '" + weeklyRptConfig.projectId + "' and SalesPeriodAndStage__c != null and ForAndnonForSaleStore__c = '\u5206\u8B72\uFF08\u4E00\u822C\uFF09' order by SalesPeriodAndStage__r.Lottery__c,SalesPeriodAndStage__c";
                        totalInfo = new SupplyInfo();
                        totalInfo.rowIndex = 59;
                        nowInfo = new SupplyInfo();
                        nowInfo.rowIndex = 58;
                        supplyedInfo = new SupplyInfo();
                        supplyedInfo.rowIndex = 57;
                        return [4, this.sfdcService.query(query).catch(function (err) {
                                throw err;
                            })];
                    case 1:
                        supplyedRes = _e.sent();
                        if (!supplyedRes || supplyedRes.totalSize === 0)
                            return [2];
                        lastSalesPeriodAndStageId = supplyedRes.records[supplyedRes.totalSize - 1].SalesPeriodAndStage__c;
                        try {
                            for (_a = tslib_1.__values(supplyedRes.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                room = _b.value;
                                this.setSupplyInfoBySobj(room, totalInfo);
                                if (room.SalesPeriodAndStage__c === lastSalesPeriodAndStageId) {
                                    this.setSupplyInfoBySobj(room, nowInfo);
                                }
                                this.setNoneForWeeklyInfo(room, this.noneInfo);
                            }
                        }
                        catch (e_23_1) { e_23 = { error: e_23_1 }; }
                        finally {
                            try {
                                if (_b && !_b.done && (_c = _a.return)) _c.call(_a);
                            }
                            finally { if (e_23) throw e_23.error; }
                        }
                        supplyedInfo.supplyedCount = totalInfo.supplyedCount - nowInfo.supplyedCount;
                        supplyedInfo.reservationCount = totalInfo.reservationCount - nowInfo.reservationCount;
                        supplyedInfo.applicationCount = totalInfo.applicationCount - nowInfo.applicationCount;
                        supplyedInfo.contractCount = totalInfo.contractCount - nowInfo.contractCount;
                        supplyedInfo.noneCount = totalInfo.noneCount - nowInfo.noneCount;
                        supplyedInfo.deliverCount = totalInfo.deliverCount - nowInfo.deliverCount;
                        optList = [totalInfo, nowInfo, supplyedInfo];
                        try {
                            for (optList_2 = tslib_1.__values(optList), optList_2_1 = optList_2.next(); !optList_2_1.done; optList_2_1 = optList_2.next()) {
                                optInfo = optList_2_1.value;
                                this.worksheet.getCell(optInfo.rowIndex, 39).value = optInfo.supplyedCount;
                                this.worksheet.getCell(optInfo.rowIndex, 41).value = optInfo.reservationCount;
                                this.worksheet.getCell(optInfo.rowIndex, 43).value = optInfo.applicationCount;
                                this.worksheet.getCell(optInfo.rowIndex, 45).value = optInfo.contractCount;
                                this.worksheet.getCell(optInfo.rowIndex, 47).value = optInfo.noneCount;
                                this.worksheet.getCell(optInfo.rowIndex, 49).value = optInfo.deliverCount;
                            }
                        }
                        catch (e_24_1) { e_24 = { error: e_24_1 }; }
                        finally {
                            try {
                                if (optList_2_1 && !optList_2_1.done && (_d = optList_2.return)) _d.call(optList_2);
                            }
                            finally { if (e_24) throw e_24.error; }
                        }
                        return [4, this.sfdcService
                                .query("SELECT count(Id) noSp FROM Room__c where Project__c = '" + weeklyRptConfig.projectId + "' and ForAndnonForSaleStore__c = '\u5206\u8B72\uFF08\u4E00\u822C\uFF09' and SalesPeriodAndStage__c = null")
                                .catch(function (err) {
                                throw err;
                            })];
                    case 2:
                        res = _e.sent();
                        if (res && res.totalSize > 0) {
                            countData = res.records[0];
                            this.worksheet.getCell("AW55").value = countData.noSp + "\u6238";
                        }
                        return [2];
                }
            });
        });
    };
    WeeklyReportService.prototype.setSupplyInfoBySobj = function (room, spInfo) {
        spInfo.supplyedCount++;
        var contact;
        var today = this.theWeek ? this.theWeek.PeriodTo : undefined;
        today = today ? today : new Date();
        if (room.Contacts__r && room.Contacts__r.totalSize > 0) {
            contact = room.Contacts__r.records[0];
            var cancellationDate = tools_service_1.Tools.strToDate(contact.CancellationDate__c);
            if (contact.ResponseKinds__c === "予約" &&
                (!contact.CancellationDate__c || (cancellationDate && cancellationDate > today))) {
                spInfo.reservationCount++;
            }
            else if (contact.ResponseKinds__c === "申込" &&
                (!contact.CancellationDate__c || (cancellationDate && cancellationDate > today))) {
                spInfo.applicationCount++;
            }
            else if (contact.ResponseKinds__c === "契約" &&
                (!contact.CancellationDate__c || (cancellationDate && cancellationDate > today))) {
                spInfo.contractCount++;
            }
            else if (contact.ResponseKinds__c === "引渡") {
                spInfo.deliverCount++;
            }
            else {
                spInfo.noneCount++;
            }
        }
        else {
            spInfo.noneCount++;
        }
    };
    WeeklyReportService.prototype.setNoneForWeeklyInfo = function (room, noneInfo) {
        if (!this.theWeek || !this.theWeek.PeriodTo)
            return;
        var contact;
        var compareDate = this.theWeek.PeriodTo;
        var Lottery = tools_service_1.Tools.strToDate(room.SalesPeriodAndStage__r.Lottery__c);
        if (!Lottery || Lottery > compareDate)
            return;
        if (room.Contacts__r && room.Contacts__r.totalSize > 0) {
            contact = room.Contacts__r.records[0];
            var cancellationDate = tools_service_1.Tools.strToDate(contact.CancellationDate__c);
            var statusList = ["予約", "申込", "契約"];
            if (statusList.includes(contact.ResponseKinds__c) && cancellationDate && cancellationDate <= compareDate) {
                noneInfo.count++;
            }
        }
        else {
            noneInfo.count++;
        }
    };
    WeeklyReportService.prototype.saveOtherData = function () {
        if (!this.worksheet)
            return;
        this.worksheet.getCell("AI43").value = this.noneInfo.count;
        this.worksheet.getCell("S43").value = this.noneInfo.count - Number(this.worksheet.getCell("AH43").value);
    };
    return WeeklyReportService;
}());
exports.WeeklyReportService = WeeklyReportService;
//# sourceMappingURL=weeklyReport.service.js.map