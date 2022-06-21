"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelService = void 0;
var tslib_1 = require("tslib");
var exceljs_1 = tslib_1.__importDefault(require("exceljs"));
var sfdc_service_1 = require("./sfdc.service");
var tools_service_1 = require("./tools.service");
var date_and_time_1 = tslib_1.__importDefault(require("date-and-time"));
var pg_service_1 = require("./pg.service");
var ExcelService = (function () {
    function ExcelService() {
        this.workbook = new exceljs_1.default.Workbook();
        this.sheetLsit = [];
        this.sfdcService = new sfdc_service_1.SfdcService();
        this.pgService = new pg_service_1.PgService();
    }
    ExcelService.prototype.addSheet = function (sheetName) {
        this.worksheet = this.workbook.addWorksheet(sheetName);
        this.sheetLsit.push(this.worksheet);
        var firstCol = this.worksheet.getColumn("A");
        firstCol.header = ["これはタイトル"];
        firstCol.width = 30;
        var tableBorder = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" }
        };
        var row = this.worksheet.getRow(2);
        row.values = ["番号", "名前", "年齢", "性別"];
        var headerCells = [];
        for (var col = 1; col <= 4; col++) {
            var cell = this.worksheet.getCell(2, col);
            cell.border = tableBorder;
            headerCells.push(cell);
        }
        this.worksheet.addRows([
            ["1", "田中", 28, "男性"],
            ["2", "斉藤", 30, "男性"],
            ["3", "山下", 25, "男性"],
            ["4", "中山", 31, "女性"]
        ], "i");
        headerCells.forEach(function (cell) {
            cell.font = {
                size: 15,
                bold: true
            };
            cell.fill = {
                type: "pattern",
                pattern: "darkTrellis",
                fgColor: { argb: "ECEFF1FC" }
            };
        });
    };
    ExcelService.prototype.getFileBuffer = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4, this.workbook.xlsx.writeBuffer()];
                    case 1: return [2, _a.sent()];
                }
            });
        });
    };
    ExcelService.prototype.getStatustable = function (fileSteam, pjcode, statustableConfig) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var _a, patternList, key, pattern, roomCellMap, roomNoSql, customfiledMap, customFiledSet, customFiledSet_1, customFiledSet_1_1, filed, customfiledSql, _b, _c, key_1, customfiledList, subSql, customfiledList_1, customfiledList_1_1, filed, sfdcService, query, result, customFiledSet_2, customFiledSet_2_1, filed, _d, _e, room, roomCell, _f, _g, filed, filedValue, filedList, index, oppResult;
            var e_1, _h, e_2, _j, e_3, _k, e_4, _l, e_5, _m, e_6, _o;
            return tslib_1.__generator(this, function (_p) {
                switch (_p.label) {
                    case 0:
                        _a = this;
                        return [4, this.workbook.xlsx.read(fileSteam)];
                    case 1:
                        _a.workbook = _p.sent();
                        this.worksheet = this.workbook.worksheets[0];
                        patternList = [
                            {
                                key: "contract_butterfly",
                                label: "契約パターン-部屋属性+蝶々",
                                rowSpacing: 5,
                                sql: "SELECT Id,Name,ProjectNames__c,BuildingCode__c,FloorPlan__c,TypeCode__c,OccupiedArea__c,SellingPrice__c,SalesPeriodAndStage__r.Name",
                                filedIndex: [
                                    { x: 1, y: 0, filedName: "FloorPlan__c" },
                                    { x: 2, y: 0, filedName: "SalesPeriodAndStage__r.Name" },
                                    { x: 0, y: 1, filedName: "TypeCode__c" },
                                    { x: 1, y: 1, filedName: "OccupiedArea__c" },
                                    { x: 2, y: 1, filedName: "SellingPrice__c" }
                                ],
                                customFiledSet: [
                                    { index: 1, relationship: "Contacts__r", apiName: "ApplicationDate__c" },
                                    { index: 2, relationship: "Contacts__r", apiName: "ReserveDate__c" },
                                    { index: 6, relationship: "Contacts__r", apiName: "ContractDate__c" },
                                    { index: 7, relationship: "Contacts__r", apiName: "DeliverDate__c" }
                                ],
                                cstPlusX: 2
                            },
                            {
                                key: "contract_info",
                                label: "契約パターン-部屋属性+個人属性",
                                rowSpacing: 8,
                                sql: "SELECT Id,Name,ProjectNames__c,BuildingCode__c,FloorPlan__c,SalesPeriodAndStage__c,SalesPeriodAndStage__r.Name,TypeCode__c,OccupiedArea__c,SellingPrice__c",
                                filedIndex: [
                                    { x: 1, y: 0, filedName: "FloorPlan__c" },
                                    { x: 2, y: 0, filedName: "SalesPeriodAndStage__r.Name" },
                                    { x: 0, y: 1, filedName: "TypeCode__c" },
                                    { x: 1, y: 1, filedName: "OccupiedArea__c" },
                                    { x: 2, y: 1, filedName: "SellingPrice__c" }
                                ],
                                cstPlusX: 2
                            },
                            {
                                key: "contract_btInfo",
                                label: "契約パターン-部屋属性+個人属性+蝶々",
                                rowSpacing: 10,
                                sql: "SELECT Id,Name,ProjectNames__c,BuildingCode__c,FloorPlan__c,TypeCode__c,OccupiedArea__c,SellingPrice__c,SalesPeriodAndStage__r.Name",
                                filedIndex: [
                                    { x: 1, y: 0, filedName: "FloorPlan__c" },
                                    { x: 2, y: 0, filedName: "SalesPeriodAndStage__r.Name" },
                                    { x: 0, y: 1, filedName: "TypeCode__c" },
                                    { x: 1, y: 1, filedName: "OccupiedArea__c" },
                                    { x: 2, y: 1, filedName: "SellingPrice__c" }
                                ],
                                customFiledSet: [
                                    { x: 8, y: 0, index: 0, relationship: "Contacts__r", apiName: "ApplicationDate__c" },
                                    { x: 9, y: 0, index: 0, relationship: "Contacts__r", apiName: "ReserveDate__c" },
                                    { x: 8, y: 1, index: 0, relationship: "Contacts__r", apiName: "ContractDate__c" },
                                    { x: 9, y: 1, index: 0, relationship: "Contacts__r", apiName: "DeliverDate__c" }
                                ],
                                cstPlusX: 2
                            },
                            {
                                key: "lottery",
                                label: "抽選パターン",
                                rowSpacing: 6,
                                sql: "SELECT Id,Name,ProjectNames__c,BuildingCode__c,FloorPlan__c,SalesPeriodAndStage__c,SalesPeriodAndStage__r.Name,TypeCode__c,OccupiedArea__c,SellingPrice__c,(SELECT Id FROM RegisterRoomNumber__r WHERE RegisterDate__c != null)",
                                filedIndex: [
                                    { x: 1, y: 0, filedName: "FloorPlan__c" },
                                    { x: 2, y: 0, filedName: "SalesPeriodAndStage__r.Name" },
                                    { x: 0, y: 1, filedName: "TypeCode__c" },
                                    { x: 1, y: 1, filedName: "OccupiedArea__c" },
                                    { x: 2, y: 1, filedName: "SellingPrice__c" }
                                ]
                            },
                            {
                                key: "request",
                                label: "要望パターン",
                                rowSpacing: 6,
                                sql: "SELECT Id,Name,ProjectNames__c,BuildingCode__c,FloorPlan__c,SalesPeriodAndStage__c,SalesPeriodAndStage__r.Name,TypeCode__c,OccupiedArea__c,PrePriceWithTax__c",
                                filedIndex: [
                                    { x: 1, y: 0, filedName: "FloorPlan__c" },
                                    { x: 2, y: 0, filedName: "SalesPeriodAndStage__r.Name" },
                                    { x: 0, y: 1, filedName: "TypeCode__c" },
                                    { x: 1, y: 1, filedName: "OccupiedArea__c" },
                                    { x: 2, y: 1, filedName: "PrePriceWithTax__c" }
                                ]
                            }
                        ];
                        key = statustableConfig.outputPattern === "contract"
                            ? statustableConfig.outputPattern + "_" + statustableConfig.contractPattern
                            : statustableConfig.outputPattern;
                        pattern = patternList.find(function (item) { return item.key === key; });
                        if (!pattern)
                            return [2];
                        roomCellMap = this.getCellsMap(pattern.rowSpacing);
                        if (roomCellMap.size == 0)
                            return [2];
                        roomNoSql = sfdc_service_1.SfdcService.getInSql(roomCellMap.keys());
                        customfiledMap = new Map();
                        customFiledSet = [];
                        if (pattern.customFiledSet) {
                            customFiledSet = pattern.customFiledSet;
                        }
                        if (statustableConfig.customFiledSet) {
                            customFiledSet = customFiledSet.concat(statustableConfig.customFiledSet);
                        }
                        if (customFiledSet) {
                            try {
                                for (customFiledSet_1 = tslib_1.__values(customFiledSet), customFiledSet_1_1 = customFiledSet_1.next(); !customFiledSet_1_1.done; customFiledSet_1_1 = customFiledSet_1.next()) {
                                    filed = customFiledSet_1_1.value;
                                    if (filed.apiName) {
                                        if (!customfiledMap.has(filed.relationship)) {
                                            customfiledMap.set(filed.relationship, []);
                                        }
                                        customfiledMap.get(filed.relationship).push(filed);
                                    }
                                }
                            }
                            catch (e_1_1) { e_1 = { error: e_1_1 }; }
                            finally {
                                try {
                                    if (customFiledSet_1_1 && !customFiledSet_1_1.done && (_h = customFiledSet_1.return)) _h.call(customFiledSet_1);
                                }
                                finally { if (e_1) throw e_1.error; }
                            }
                        }
                        customfiledSql = "";
                        if (customfiledMap.size > 0) {
                            try {
                                for (_b = tslib_1.__values(customfiledMap.keys()), _c = _b.next(); !_c.done; _c = _b.next()) {
                                    key_1 = _c.value;
                                    customfiledList = customfiledMap.get(key_1);
                                    subSql = "";
                                    try {
                                        for (customfiledList_1 = (e_3 = void 0, tslib_1.__values(customfiledList)), customfiledList_1_1 = customfiledList_1.next(); !customfiledList_1_1.done; customfiledList_1_1 = customfiledList_1.next()) {
                                            filed = customfiledList_1_1.value;
                                            subSql += filed.apiName + ",";
                                        }
                                    }
                                    catch (e_3_1) { e_3 = { error: e_3_1 }; }
                                    finally {
                                        try {
                                            if (customfiledList_1_1 && !customfiledList_1_1.done && (_k = customfiledList_1.return)) _k.call(customfiledList_1);
                                        }
                                        finally { if (e_3) throw e_3.error; }
                                    }
                                    subSql = subSql.substring(0, subSql.length - 1);
                                    customfiledSql += ",(SELECT " + subSql + " FROM " + key_1 + " limit 1)";
                                }
                            }
                            catch (e_2_1) { e_2 = { error: e_2_1 }; }
                            finally {
                                try {
                                    if (_c && !_c.done && (_j = _b.return)) _j.call(_b);
                                }
                                finally { if (e_2) throw e_2.error; }
                            }
                        }
                        sfdcService = new sfdc_service_1.SfdcService();
                        query = "" + pattern.sql + customfiledSql + " FROM Room__c WHERE Project__c = '" + pjcode + "' AND Name IN " + roomNoSql;
                        return [4, sfdcService.query(query)];
                    case 2:
                        result = _p.sent();
                        if (!result || result.totalSize === 0)
                            return [2];
                        this.worksheet.getCell(4, 3).value = "\u7269\u4EF6\u540D  " + result.records[0].ProjectNames__c;
                        if (pattern.cstPlusX) {
                            try {
                                for (customFiledSet_2 = tslib_1.__values(customFiledSet), customFiledSet_2_1 = customFiledSet_2.next(); !customFiledSet_2_1.done; customFiledSet_2_1 = customFiledSet_2.next()) {
                                    filed = customFiledSet_2_1.value;
                                    pattern.filedIndex.push({
                                        x: filed.x ? filed.x : pattern.cstPlusX + (filed.index % 5),
                                        y: filed.y ? filed.y : filed.index <= 5 ? 0 : 1,
                                        filedName: filed.relationship + "." + filed.apiName
                                    });
                                }
                            }
                            catch (e_4_1) { e_4 = { error: e_4_1 }; }
                            finally {
                                try {
                                    if (customFiledSet_2_1 && !customFiledSet_2_1.done && (_l = customFiledSet_2.return)) _l.call(customFiledSet_2);
                                }
                                finally { if (e_4) throw e_4.error; }
                            }
                        }
                        try {
                            for (_d = tslib_1.__values(result.records), _e = _d.next(); !_e.done; _e = _d.next()) {
                                room = _e.value;
                                roomCell = roomCellMap.get(room.Name);
                                if (roomCell) {
                                    try {
                                        for (_f = (e_6 = void 0, tslib_1.__values(pattern.filedIndex)), _g = _f.next(); !_g.done; _g = _f.next()) {
                                            filed = _g.value;
                                            filedValue = void 0;
                                            if (filed.filedName.indexOf(".") !== -1) {
                                                filedList = filed.filedName.split(".");
                                                filedValue = room;
                                                for (index = 0; index < filedList.length; index++) {
                                                    if (!filedValue)
                                                        continue;
                                                    filedValue = filedValue[filedList[index]];
                                                    if (index === 0 && filedValue) {
                                                        if ("totalSize" in filedValue) {
                                                            if (filedValue && filedValue.totalSize > 0) {
                                                                filedValue = filedValue.records[0];
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            else {
                                                filedValue = room[filed.filedName];
                                            }
                                            this.worksheet.getCell(roomCell.cell.row + filed.x, roomCell.cell.col + filed.y * roomCell.mergeColNum).value = filedValue;
                                        }
                                    }
                                    catch (e_6_1) { e_6 = { error: e_6_1 }; }
                                    finally {
                                        try {
                                            if (_g && !_g.done && (_o = _f.return)) _o.call(_f);
                                        }
                                        finally { if (e_6) throw e_6.error; }
                                    }
                                    if (pattern.key === "lottery") {
                                        oppResult = room.RegisterRoomNumber__r;
                                        if (oppResult && oppResult.totalSize > 0) {
                                            this.worksheet.getCell(roomCell.cell.row + 3, roomCell.cell.col + 0).value = oppResult.totalSize;
                                        }
                                    }
                                }
                            }
                        }
                        catch (e_5_1) { e_5 = { error: e_5_1 }; }
                        finally {
                            try {
                                if (_e && !_e.done && (_m = _d.return)) _m.call(_d);
                            }
                            finally { if (e_5) throw e_5.error; }
                        }
                        return [2];
                }
            });
        });
    };
    ExcelService.prototype.getStatustable2 = function (filePath, pjcode) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var _a, colA, rowMap, StageName, roomCellMap, roomNoSql, sfdcService, query, result, _b, _c, room, roomCell, contactsResult, contact;
            var e_7, _d;
            return tslib_1.__generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        _a = this;
                        return [4, this.workbook.xlsx.readFile(filePath)];
                    case 1:
                        _a.workbook = _e.sent();
                        this.worksheet = this.workbook.worksheets[0];
                        colA = this.worksheet.getColumn("A");
                        rowMap = new Map();
                        StageName = "";
                        colA.eachCell(function (cell, rowNumber) {
                            if (cell.value !== null && StageName !== cell.value) {
                                StageName = cell.value ? cell.value.toString() : "";
                                rowMap.set(StageName, rowNumber);
                            }
                        });
                        roomCellMap = this.getCellsMap(5);
                        if (roomCellMap.size == 0)
                            return [2];
                        roomNoSql = sfdc_service_1.SfdcService.getInSql(roomCellMap.keys());
                        sfdcService = new sfdc_service_1.SfdcService();
                        query = "SELECT Id,Name,ProjectNames__c,BuildingCode__c,FloorPlan__c,TypeCode__c,OccupiedArea__c,SellingPrice__c,SalesPeriodAndStage__r.Name,(SELECT ApplicationDate__c,ReserveDate__c,ContractDate__c,DeliverDate__c FROM Contacts__r ORDER BY ApplicationDate__c DESC limit 1) FROM Room__c WHERE ProjectCode__c = '" + pjcode + "' AND Name IN " + roomNoSql;
                        return [4, sfdcService.query(query)];
                    case 2:
                        result = _e.sent();
                        if (!result || result.totalSize === 0)
                            return [2];
                        this.worksheet.getCell(4, 3).value = "\u7269\u4EF6\u540D  " + result.records[0].ProjectNames__c;
                        try {
                            for (_b = tslib_1.__values(result.records), _c = _b.next(); !_c.done; _c = _b.next()) {
                                room = _c.value;
                                roomCell = roomCellMap.get(room.Name);
                                if (roomCell) {
                                    this.worksheet.getCell(roomCell.cell.row + 1, roomCell.cell.col).value = room.FloorPlan__c;
                                    if (room.SalesPeriodAndStage__r) {
                                        this.worksheet.getCell(roomCell.cell.row + 2, roomCell.cell.col).value = room.SalesPeriodAndStage__r.Name;
                                    }
                                    this.worksheet.getCell(roomCell.cell.row, roomCell.cell.col + 2).value = room.TypeCode__c;
                                    this.worksheet.getCell(roomCell.cell.row + 1, roomCell.cell.col + 2).value = room.OccupiedArea__c;
                                    this.worksheet.getCell(roomCell.cell.row + 2, roomCell.cell.col + 2).value = room.SellingPrice__c;
                                    contactsResult = room.Contacts__r;
                                    if (contactsResult && contactsResult.totalSize > 0) {
                                        contact = contactsResult.records[0];
                                        this.worksheet.getCell(roomCell.cell.row + 3, roomCell.cell.col).value = contact.ApplicationDate__c;
                                        this.worksheet.getCell(roomCell.cell.row + 4, roomCell.cell.col).value = contact.ReserveDate__c;
                                        this.worksheet.getCell(roomCell.cell.row + 3, roomCell.cell.col + 2).value = contact.ContractDate__c;
                                        this.worksheet.getCell(roomCell.cell.row + 4, roomCell.cell.col + 2).value = contact.DeliverDate__c;
                                    }
                                }
                            }
                        }
                        catch (e_7_1) { e_7 = { error: e_7_1 }; }
                        finally {
                            try {
                                if (_c && !_c.done && (_d = _b.return)) _d.call(_b);
                            }
                            finally { if (e_7) throw e_7.error; }
                        }
                        return [2];
                }
            });
        });
    };
    ExcelService.prototype.getCellsMap = function (rowSpacing) {
        var roomCellMap = new Map();
        if (!this.worksheet)
            return roomCellMap;
        var maxRow = this.worksheet.rowCount;
        var startRowNum = 12;
        var _loop_1 = function () {
            var masterCell;
            var targetRow = this_1.worksheet.getRow(startRowNum);
            if (targetRow.actualCellCount == 0) {
                startRowNum++;
                return "continue";
            }
            targetRow.eachCell(function (cell, colNumber) {
                if (colNumber == 1)
                    return;
                if (cell.value && (!masterCell || masterCell.cell.value !== cell.value)) {
                    masterCell = { cell: cell, mergeColNum: 1 };
                    roomCellMap.set(cell.value.toString(), masterCell);
                }
                else {
                    masterCell.mergeColNum++;
                }
            });
            startRowNum += rowSpacing;
        };
        var this_1 = this;
        while (startRowNum <= maxRow) {
            _loop_1();
        }
        return roomCellMap;
    };
    ExcelService.prototype.gennaretaWeeklyReport = function (weeklyRptConfig) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var _a, fromDate, toDate, queryStartDate, queryEndQuery, _b, _c, item, promiseJobList;
            var e_8, _d;
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
                        catch (e_8_1) { e_8 = { error: e_8_1 }; }
                        finally {
                            try {
                                if (_c && !_c.done && (_d = _b.return)) _d.call(_b);
                            }
                            finally { if (e_8) throw e_8.error; }
                        }
                        promiseJobList = [];
                        promiseJobList.push(this.setTitleAndAvPrice(weeklyRptConfig));
                        promiseJobList.push(this.setSalesPeriodMaster());
                        promiseJobList.push(this.setWeeklyInfo(weeklyRptConfig, queryStartDate, queryEndQuery));
                        promiseJobList.push(this.setBrilliaClub(weeklyRptConfig));
                        promiseJobList.push(this.setSubQuestion(weeklyRptConfig));
                        promiseJobList.push(this.setOOSInfo(weeklyRptConfig));
                        return [4, Promise.all(promiseJobList).catch(function (err) {
                                throw err;
                            })];
                    case 2:
                        _e.sent();
                        return [2];
                }
            });
        });
    };
    ExcelService.prototype.setTitleAndAvPrice = function (weeklyRptConfig) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var accountRes, acc, _a, _b, build, val, _c, _d, sellerInfo, _e, _f, saleInfo, buildIdList, _g, _h, build, buildIdSql, roomRes, avInfo, buildAvMap, _j, _k, room, buildAvInfo, bulidSellText, bulidFormalText, bulidPlayText, avPrice, _l, _m, buildId, avPriceInfo;
            var e_9, _o, e_10, _p, e_11, _q, e_12, _r, e_13, _s, e_14, _t;
            return tslib_1.__generator(this, function (_u) {
                switch (_u.label) {
                    case 0:
                        if (!this.worksheet)
                            return [2];
                        return [4, this.sfdcService.query("SELECT Id,Name,LocationLotNumber__c,(SELECT Id,OldeBrilliaTran__c FROM Buildings__r),(SELECT Id,CorporationName__c,Share__c FROM Project_SellerInfo__r order by SortOrder__c),(SELECT Id,CorporationName__c,Share__c FROM Project_SaleInfo__r order by SortOrder__c) FROM Account WHERE Id = '" + weeklyRptConfig.projectId + "'")];
                    case 1:
                        accountRes = _u.sent();
                        acc = accountRes.records[0];
                        this.worksheet.getCell("D2").value = acc.Name;
                        this.worksheet.getCell("C3").value = weeklyRptConfig.selectWeek.Weeks__c;
                        this.worksheet.getCell("C6").value = acc.LocationLotNumber__c;
                        try {
                            for (_a = tslib_1.__values(acc.Buildings__r.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                build = _b.value;
                                if (build.OldeBrilliaTran__c) {
                                    this.worksheet.getCell("C8").value = build.OldeBrilliaTran__c;
                                    break;
                                }
                            }
                        }
                        catch (e_9_1) { e_9 = { error: e_9_1 }; }
                        finally {
                            try {
                                if (_b && !_b.done && (_o = _a.return)) _o.call(_a);
                            }
                            finally { if (e_9) throw e_9.error; }
                        }
                        val = "";
                        try {
                            for (_c = tslib_1.__values(acc.Project_SellerInfo__r.records), _d = _c.next(); !_d.done; _d = _c.next()) {
                                sellerInfo = _d.value;
                                if (sellerInfo.CorporationName__c) {
                                    val += sellerInfo.CorporationName__c + "(" + sellerInfo.Share__c + ")\u3001";
                                }
                            }
                        }
                        catch (e_10_1) { e_10 = { error: e_10_1 }; }
                        finally {
                            try {
                                if (_d && !_d.done && (_p = _c.return)) _p.call(_c);
                            }
                            finally { if (e_10) throw e_10.error; }
                        }
                        val = val.length > 0 ? val.substring(0, val.length - 1) : "";
                        this.worksheet.getCell("C11").value = val;
                        val = "";
                        try {
                            for (_e = tslib_1.__values(acc.Project_SaleInfo__r.records), _f = _e.next(); !_f.done; _f = _e.next()) {
                                saleInfo = _f.value;
                                if (saleInfo.CorporationName__c) {
                                    val += saleInfo.CorporationName__c + "(" + saleInfo.Share__c + ")\u3001";
                                }
                            }
                        }
                        catch (e_11_1) { e_11 = { error: e_11_1 }; }
                        finally {
                            try {
                                if (_f && !_f.done && (_q = _e.return)) _q.call(_e);
                            }
                            finally { if (e_11) throw e_11.error; }
                        }
                        val = val.length > 0 ? val.substring(0, val.length - 1) : "";
                        this.worksheet.getCell("C13").value = val;
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
                        catch (e_12_1) { e_12 = { error: e_12_1 }; }
                        finally {
                            try {
                                if (_h && !_h.done && (_r = _g.return)) _r.call(_g);
                            }
                            finally { if (e_12) throw e_12.error; }
                        }
                        buildIdSql = sfdc_service_1.SfdcService.getInSql(buildIdList);
                        if (!(buildIdList.length > 0)) return [3, 3];
                        return [4, this.sfdcService.query("SELECT Id,Name,SellingPrice__c,PrePriceWithTax__c,FullPriceFlag__c,ForAndnonForSaleStore__c,Building__c,BuildingName__c,OccupiedArea__c FROM Room__c WHERE Building__c IN " + buildIdSql)];
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
                                        avInfo.sellTotalPrice += room.SellingPrice__c ? room.SellingPrice__c : 0;
                                        avInfo.sellArea += room.OccupiedArea__c ? room.OccupiedArea__c : 0;
                                        buildAvInfo.sellNum++;
                                        buildAvInfo.sellTotalPrice += room.SellingPrice__c ? room.SellingPrice__c : 0;
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
                            catch (e_13_1) { e_13 = { error: e_13_1 }; }
                            finally {
                                try {
                                    if (_k && !_k.done && (_s = _j.return)) _s.call(_j);
                                }
                                finally { if (e_13) throw e_13.error; }
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
                            catch (e_14_1) { e_14 = { error: e_14_1 }; }
                            finally {
                                try {
                                    if (_m && !_m.done && (_t = _l.return)) _t.call(_l);
                                }
                                finally { if (e_14) throw e_14.error; }
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
    ExcelService.prototype.initWeekDate = function (theDate, xindex, weekName) {
        return {
            date: theDate,
            xindex: xindex,
            weekName: weekName,
            InquiryDateFirst__c: 0,
            OSSRelationLatestDate__c: 0,
            VisitDate__c: 0,
            ReturnVisitDate__c: 0,
            ReturnLatestVisitDate__c: 0,
            RegisterDate__c: 0,
            RegisterCancelDate__c: 0,
            ReserveDate__c: 0,
            ReserveCancell: 0,
            ApplicationDate__c: 0,
            ApplicationCancell: 0,
            ContractDate__c: 0,
            ContractCancell: 0,
            DeliverDate__c: 0
        };
    };
    ExcelService.prototype.setWeeklyInfo = function (weeklyRptConfig, queryStartDate, queryEndQuery) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var weekInfo, startDate, endDate, weekDateMap, index, theDate, index, theDate, startDateTime, endDateTime, oppQuery, oppList, _a, _b, opp, contractQuery, contractList, _c, _d, contract, _e, _f, weekDate, pgRes, otherOpp, visitCount, theWeekDate, InquiryVisitNum;
            var e_15, _g, e_16, _h, e_17, _j;
            return tslib_1.__generator(this, function (_k) {
                switch (_k.label) {
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
                        startDateTime = queryStartDate + "T00:00:00Z";
                        endDateTime = queryEndQuery + "T00:00:00Z";
                        oppQuery = "SELECT Id,InquiryDateFirst__c,OSSRelationLatestDate__c,VisitDate__c,ReturnVisitDate__c,ReturnLatestVisitDate__c,RegisterDate__c,RegisterCancelDate__c FROM Opportunity WHERE AccountId = '" + weeklyRptConfig.projectId + "' AND ((InquiryDateFirst__c >= " + startDateTime + " AND InquiryDateFirst__c <= " + endDateTime + ") OR (OSSRelationLatestDate__c >= " + startDateTime + " AND OSSRelationLatestDate__c <= " + endDateTime + ") OR (VisitDate__c >= " + startDateTime + " AND VisitDate__c <= " + endDateTime + ") OR (ReturnVisitDate__c >= " + startDateTime + " AND ReturnVisitDate__c <= " + endDateTime + ") OR (ReturnLatestVisitDate__c >= " + startDateTime + " AND ReturnLatestVisitDate__c <= " + endDateTime + ") OR (RegisterDate__c >= " + queryStartDate + " AND RegisterDate__c <= " + queryEndQuery + ") OR (RegisterCancelDate__c >= " + queryStartDate + " AND RegisterCancelDate__c <= " + queryEndQuery + "))";
                        return [4, this.sfdcService.query(oppQuery)];
                    case 1:
                        oppList = _k.sent();
                        if (oppList.totalSize > 0) {
                            try {
                                for (_a = tslib_1.__values(oppList.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                    opp = _b.value;
                                    this.setWeekTableInfo(opp, weekDateMap, "InquiryDateFirst__c");
                                    this.setWeekTableInfo(opp, weekDateMap, "OSSRelationLatestDate__c");
                                    this.setWeekTableInfo(opp, weekDateMap, "VisitDate__c");
                                    this.setWeekTableInfo(opp, weekDateMap, "ReturnVisitDate__c");
                                    this.setWeekTableInfo(opp, weekDateMap, "ReturnLatestVisitDate__c");
                                    this.setWeekTableInfo(opp, weekDateMap, "RegisterDate__c");
                                    this.setWeekTableInfo(opp, weekDateMap, "RegisterCancelDate__c");
                                }
                            }
                            catch (e_15_1) { e_15 = { error: e_15_1 }; }
                            finally {
                                try {
                                    if (_b && !_b.done && (_g = _a.return)) _g.call(_a);
                                }
                                finally { if (e_15) throw e_15.error; }
                            }
                        }
                        contractQuery = "SELECT Id,ContractDate__c,ApplicationDate__c,CancellationDate__c,DeliverDate__c,ReserveDate__c,ResponseKinds__c FROM Contract WHERE AccountId = '" + weeklyRptConfig.projectId + "'";
                        return [4, this.sfdcService.query(contractQuery)];
                    case 2:
                        contractList = _k.sent();
                        if (contractList.totalSize > 0) {
                            try {
                                for (_c = tslib_1.__values(contractList.records), _d = _c.next(); !_d.done; _d = _c.next()) {
                                    contract = _d.value;
                                    if (contract.ResponseKinds__c === "予約") {
                                        this.setWeekTableInfo(contract, weekDateMap, "ReserveDate__c");
                                        this.setWeekTableInfo(contract, weekDateMap, "CancellationDate__c", "ReserveCancell");
                                    }
                                    else if (contract.ResponseKinds__c === "申込") {
                                        this.setWeekTableInfo(contract, weekDateMap, "ApplicationDate__c");
                                        this.setWeekTableInfo(contract, weekDateMap, "CancellationDate__c", "ApplicationCancell");
                                    }
                                    else if (contract.ResponseKinds__c === "契約") {
                                        this.setWeekTableInfo(contract, weekDateMap, "ContractDate__c");
                                        this.setWeekTableInfo(contract, weekDateMap, "CancellationDate__c", "ContractCancell");
                                    }
                                    else if (contract.ResponseKinds__c === "引渡") {
                                        this.setWeekTableInfo(contract, weekDateMap, "DeliverDate__c");
                                    }
                                }
                            }
                            catch (e_16_1) { e_16 = { error: e_16_1 }; }
                            finally {
                                try {
                                    if (_d && !_d.done && (_h = _c.return)) _h.call(_c);
                                }
                                finally { if (e_16) throw e_16.error; }
                            }
                        }
                        try {
                            for (_e = tslib_1.__values(weekDateMap.values()), _f = _e.next(); !_f.done; _f = _e.next()) {
                                weekDate = _f.value;
                                this.worksheet.getCell(34, weekDate.xindex).value = weekDate.InquiryDateFirst__c;
                                this.worksheet.getCell(35, weekDate.xindex).value = weekDate.OSSRelationLatestDate__c;
                                this.worksheet.getCell(36, weekDate.xindex).value = weekDate.VisitDate__c;
                                this.worksheet.getCell(37, weekDate.xindex).value = weekDate.ReturnVisitDate__c;
                                this.worksheet.getCell(38, weekDate.xindex).value = weekDate.ReturnLatestVisitDate__c;
                                if (weekDate.date) {
                                    this.worksheet.getCell(39, weekDate.xindex).value = weekDate.RegisterDate__c;
                                    this.worksheet.getCell(39, weekDate.xindex + 1).value = -weekDate.RegisterCancelDate__c;
                                    this.worksheet.getCell(40, weekDate.xindex).value = weekDate.ReserveDate__c;
                                    this.worksheet.getCell(40, weekDate.xindex + 1).value = -weekDate.ReserveCancell;
                                    this.worksheet.getCell(41, weekDate.xindex).value = weekDate.ApplicationDate__c;
                                    this.worksheet.getCell(41, weekDate.xindex + 1).value = -weekDate.ApplicationCancell;
                                    this.worksheet.getCell(42, weekDate.xindex).value = weekDate.ContractDate__c;
                                    this.worksheet.getCell(42, weekDate.xindex + 1).value = -weekDate.ContractCancell;
                                }
                                else {
                                    this.worksheet.getCell(39, weekDate.xindex).value = weekDate.RegisterDate__c - weekDate.RegisterCancelDate__c;
                                    this.worksheet.getCell(40, weekDate.xindex).value = weekDate.ReserveDate__c - weekDate.ReserveCancell;
                                    this.worksheet.getCell(41, weekDate.xindex).value = weekDate.ApplicationDate__c - weekDate.ApplicationCancell;
                                    this.worksheet.getCell(42, weekDate.xindex).value = weekDate.ContractDate__c - weekDate.ContractCancell;
                                }
                                this.worksheet.getCell(44, weekDate.xindex).value = weekDate.DeliverDate__c;
                            }
                        }
                        catch (e_17_1) { e_17 = { error: e_17_1 }; }
                        finally {
                            try {
                                if (_f && !_f.done && (_j = _e.return)) _j.call(_e);
                            }
                            finally { if (e_17) throw e_17.error; }
                        }
                        return [4, this.pgService
                                .query("\n          SELECT\n            accountid as \"AccountId\",\n            array_to_string(array(SELECT count(1) FROM sfdc.opportunity where InquiryDateFirst__c is not null and accountid = t1.accountid),',') as \"InquiryDateFirst__c\",\n            array_to_string(array(SELECT count(1) FROM sfdc.opportunity where OSSRelationLatestDate__c is not null and accountid = t1.accountid),',') as \"OSSRelationLatestDate__c\",\n            array_to_string(array(SELECT count(1) FROM sfdc.opportunity where VisitDate__c is not null and accountid = t1.accountid),',') as \"VisitDate__c\",\n            array_to_string(array(SELECT count(1) FROM sfdc.opportunity where ReturnVisitDate__c is not null and accountid = t1.accountid),',') as \"ReturnVisitDate__c\",\n            array_to_string(array(SELECT count(1) FROM sfdc.opportunity where ReturnLatestVisitDate__c is not null and accountid = t1.accountid),',') as \"ReturnLatestVisitDate__c\",\n            array_to_string(array(SELECT count(1) FROM sfdc.opportunity where RegisterDate__c is not null and accountid = t1.accountid),',') as \"RegisterDate__c\",\n            array_to_string(array(SELECT count(1) FROM sfdc.opportunity where RegisterCancelDate__c is not null and accountid = t1.accountid),',') as \"RegisterCancelDate__c\",\n            array_to_string(array(SELECT count(1) FROM sfdc.opportunity where InquiryDateFirst__c is not null and VisitDate__c is not null AND accountid = t1.accountid),',') as \"InquiryVisitNum\"\n          FROM sfdc.opportunity t1 where t1.accountid = '" + weeklyRptConfig.projectId + "' group by accountid")
                                .catch(function (error) {
                                throw error;
                            })];
                    case 3:
                        pgRes = _k.sent();
                        if (pgRes && pgRes.rowCount > 0) {
                            otherOpp = pgRes.rows[0];
                            this.worksheet.getCell("AI34").value = Number(otherOpp.InquiryDateFirst__c);
                            this.worksheet.getCell("AI35").value = Number(otherOpp.OSSRelationLatestDate__c);
                            this.worksheet.getCell("AI36").value = Number(otherOpp.VisitDate__c);
                            this.worksheet.getCell("AI37").value = Number(otherOpp.ReturnVisitDate__c);
                            this.worksheet.getCell("AI38").value = Number(otherOpp.ReturnLatestVisitDate__c);
                            this.worksheet.getCell("AI39").value = Number(otherOpp.RegisterDate__c) - Number(otherOpp.RegisterCancelDate__c);
                            this.worksheet.getCell("S34").value =
                                Number(otherOpp.InquiryDateFirst__c) - Number(this.worksheet.getCell("AH34").value);
                            this.worksheet.getCell("S35").value =
                                Number(otherOpp.OSSRelationLatestDate__c) - Number(this.worksheet.getCell("AH35").value);
                            this.worksheet.getCell("S36").value = Number(otherOpp.VisitDate__c) - Number(this.worksheet.getCell("AH36").value);
                            this.worksheet.getCell("S37").value =
                                Number(otherOpp.ReturnVisitDate__c) - Number(this.worksheet.getCell("AH37").value);
                            this.worksheet.getCell("S38").value =
                                Number(otherOpp.ReturnLatestVisitDate__c) - Number(this.worksheet.getCell("AH38").value);
                            this.worksheet.getCell("S39").value =
                                Number(otherOpp.RegisterDate__c) -
                                    Number(otherOpp.RegisterCancelDate__c) -
                                    Number(this.worksheet.getCell("AH39").value);
                            visitCount = Number(otherOpp.InquiryDateFirst__c) +
                                Number(otherOpp.ReturnVisitDate__c) +
                                Number(otherOpp.ReturnLatestVisitDate__c);
                            this.worksheet.getCell("AL33").value = Number(otherOpp.InquiryDateFirst__c) / weeklyRptConfig.selectWeek.weekNum;
                            this.worksheet.getCell("AL34").value = visitCount / weeklyRptConfig.selectWeek.weekNum;
                            this.worksheet.getCell("AL35").value = Number(otherOpp.ReturnVisitDate__c) / visitCount;
                            this.worksheet.getCell("AL36").value = Number(otherOpp.ReturnLatestVisitDate__c) / visitCount;
                            this.worksheet.getCell("AL37").value = Number(otherOpp.InquiryVisitNum);
                            this.worksheet.getCell("AL38").value = Number(otherOpp.InquiryVisitNum) / visitCount;
                            this.worksheet.getCell("AL39").value = Number(otherOpp.InquiryVisitNum) / Number(otherOpp.VisitDate__c);
                            theWeekDate = weekDateMap.get("今週");
                            InquiryVisitNum = theWeekDate && theWeekDate.InquiryVisitNum ? theWeekDate.InquiryVisitNum : 0;
                            this.worksheet.getCell("AL40").value = InquiryVisitNum;
                            this.worksheet.getCell("AL41").value = InquiryVisitNum / Number(otherOpp.VisitDate__c);
                        }
                        return [2];
                }
            });
        });
    };
    ExcelService.prototype.setSalesPeriodMaster = function () {
        var _a, _b, _c, _d;
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var query, result, salesPeriodGroupIdList, salesPeriodMasterId, salesPeriodMap, _e, _f, spm, query2, result2, _g, _h, aggRes, spTgt, spTgt, rowIndex, _j, _k, spInfo;
            var e_18, _l, e_19, _m, e_20, _o;
            return tslib_1.__generator(this, function (_p) {
                switch (_p.label) {
                    case 0:
                        if (!this.worksheet)
                            return [2];
                        query = "SELECT Id,Name,SalesPeriodGroupMaster__c,SalesPeriodGroupMaster__r.Name,RegisterPeriodStart__c,RegisterPeriodEnd__c,Lottery__c,ContractDate__c from SalesPeriodMaster__c where UsedFlg__c = true ORDER BY RegisterPeriodStart__c,Lottery__c,ContractDate__c,SalesPeriodGroupMaster__c";
                        return [4, this.sfdcService.query(query)];
                    case 1:
                        result = _p.sent();
                        salesPeriodGroupIdList = [];
                        salesPeriodMasterId = [];
                        salesPeriodMap = new Map();
                        try {
                            for (_e = tslib_1.__values(result.records), _f = _e.next(); !_f.done; _f = _e.next()) {
                                spm = _f.value;
                                if (spm.SalesPeriodGroupMaster__c) {
                                    if (salesPeriodGroupIdList.includes(spm.SalesPeriodGroupMaster__c)) {
                                        continue;
                                    }
                                    salesPeriodGroupIdList.push(spm.SalesPeriodGroupMaster__c);
                                    salesPeriodMap.set(spm.SalesPeriodGroupMaster__c, spm);
                                    spm.RoomTotal = 0;
                                    spm.Name = spm.SalesPeriodGroupMaster__r ? spm.SalesPeriodGroupMaster__r.Name : spm.Name;
                                }
                                else {
                                    salesPeriodMasterId.push(spm.Id);
                                    salesPeriodMap.set(spm.Id, spm);
                                }
                            }
                        }
                        catch (e_18_1) { e_18 = { error: e_18_1 }; }
                        finally {
                            try {
                                if (_f && !_f.done && (_l = _e.return)) _l.call(_e);
                            }
                            finally { if (e_18) throw e_18.error; }
                        }
                        query2 = "SELECT SalesPeriodAndStage__r.SalesPeriodGroupMaster__c SalesPeriodGroupMaster__c,SalesPeriodAndStage__c,count(Id) RoomTotal FROM Room__c WHERE SalesPeriodAndStage__c IN " + sfdc_service_1.SfdcService.getInSql(salesPeriodMasterId) + " OR SalesPeriodAndStage__r.SalesPeriodGroupMaster__c IN " + sfdc_service_1.SfdcService.getInSql(salesPeriodGroupIdList) + " GROUP BY SalesPeriodAndStage__r.SalesPeriodGroupMaster__c,SalesPeriodAndStage__c";
                        return [4, this.sfdcService.query(query2)];
                    case 2:
                        result2 = _p.sent();
                        try {
                            for (_g = tslib_1.__values(result2.records), _h = _g.next(); !_h.done; _h = _g.next()) {
                                aggRes = _h.value;
                                if (aggRes.SalesPeriodGroupMaster__c) {
                                    spTgt = salesPeriodMap.get(aggRes.SalesPeriodGroupMaster__c);
                                    if (spTgt)
                                        spTgt.RoomTotal += Number(aggRes.RoomTotal);
                                }
                                else {
                                    spTgt = salesPeriodMap.get(aggRes.SalesPeriodAndStage__c);
                                    if (spTgt)
                                        spTgt.RoomTotal = Number(aggRes.RoomTotal);
                                }
                            }
                        }
                        catch (e_19_1) { e_19 = { error: e_19_1 }; }
                        finally {
                            try {
                                if (_h && !_h.done && (_m = _g.return)) _m.call(_g);
                            }
                            finally { if (e_19) throw e_19.error; }
                        }
                        rowIndex = 3;
                        try {
                            for (_j = tslib_1.__values(salesPeriodMap.values()), _k = _j.next(); !_k.done; _k = _j.next()) {
                                spInfo = _k.value;
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
                        catch (e_20_1) { e_20 = { error: e_20_1 }; }
                        finally {
                            try {
                                if (_k && !_k.done && (_o = _j.return)) _o.call(_j);
                            }
                            finally { if (e_20) throw e_20.error; }
                        }
                        return [2];
                }
            });
        });
    };
    ExcelService.prototype.setBrilliaClub = function (WeeklyRptConfig) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var startDate, endDate, bcQuery, bcList, webInfo, newsLetterInfo, _a, _b, opp, optInfo, JoinBCDate, VisitDate, rdtRes, BCRecordTypeId, bcRes, BCTotalInfo;
            var e_21, _c;
            return tslib_1.__generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        if (!this.theWeek || !this.worksheet)
                            return [2];
                        startDate = this.theWeek.PeriodFrom__c + "T00:00:00Z";
                        endDate = this.theWeek.PeriodTo__c + "T00:00:00Z";
                        bcQuery = "SELECT Id,RecordTypeId,VisitDate__c,JoinBCDate__c,NewsletterAppRoute__c,ENewsletterProperty__c,NewsletterProperty__c,AccountId FROM Opportunity WHERE AccountId = '" + WeeklyRptConfig.projectId + "' AND RecordType.Name = 'BC' AND VisitDate__c >= " + startDate + " AND VisitDate__c <= " + endDate;
                        return [4, this.sfdcService.query(bcQuery)];
                    case 1:
                        bcList = _d.sent();
                        webInfo = {
                            visitNum: 0,
                            joinedNum: 0,
                            liveJoinNum: 0,
                            get notJoinNum() {
                                return this.visitNum - this.joinedNum;
                            }
                        };
                        newsLetterInfo = {
                            visitNum: 0,
                            joinedNum: 0,
                            liveJoinNum: 0,
                            get notJoinNum() {
                                return this.visitNum - this.joinedNum;
                            }
                        };
                        if (bcList.totalSize > 0) {
                            try {
                                for (_a = tslib_1.__values(bcList.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                    opp = _b.value;
                                    optInfo = void 0;
                                    if (opp.ENewsletterProperty__c === opp.AccountId) {
                                        optInfo = webInfo;
                                    }
                                    if (opp.NewsletterProperty__c === opp.AccountId) {
                                        optInfo = newsLetterInfo;
                                    }
                                    if (!optInfo)
                                        continue;
                                    if (opp.VisitDate__c)
                                        optInfo.visitNum++;
                                    JoinBCDate = tools_service_1.Tools.strToDate(opp.JoinBCDate__c);
                                    VisitDate = tools_service_1.Tools.strToDate(opp.VisitDate__c.substring(0, 10));
                                    if (JoinBCDate && VisitDate && JoinBCDate < VisitDate) {
                                        optInfo.joinedNum++;
                                    }
                                    else if (opp.NewsletterAppRoute__c === "販売センター" &&
                                        opp.JoinBCDate__c &&
                                        opp.JoinBCDate__c === opp.VisitDate__c.substring(0, 10)) {
                                        optInfo.liveJoinNum++;
                                    }
                                }
                            }
                            catch (e_21_1) { e_21 = { error: e_21_1 }; }
                            finally {
                                try {
                                    if (_b && !_b.done && (_c = _a.return)) _c.call(_a);
                                }
                                finally { if (e_21) throw e_21.error; }
                            }
                        }
                        this.worksheet.getCell("C49").value = webInfo.notJoinNum;
                        this.worksheet.getCell("E49").value = webInfo.liveJoinNum;
                        this.worksheet.getCell("G49").value = webInfo.liveJoinNum / webInfo.notJoinNum;
                        this.worksheet.getCell("C50").value = newsLetterInfo.notJoinNum;
                        this.worksheet.getCell("E50").value = newsLetterInfo.liveJoinNum;
                        this.worksheet.getCell("G50").value = newsLetterInfo.liveJoinNum / newsLetterInfo.notJoinNum;
                        return [4, this.sfdcService.getRecordType("Opportunity", "BC")];
                    case 2:
                        rdtRes = _d.sent();
                        BCRecordTypeId = rdtRes.records[0].Id;
                        return [4, this.pgService.query("\n    SELECT\n      t1.accountid as \"AccountId\",\n      array_to_string(array(SELECT count(1) FROM sfdc.opportunity where recordtypeid = t1.recordtypeid and accountid = t1.accountid and ENewsletterProperty__c = t1.accountid and VisitDate__c <= '" + this.theWeek.PeriodTo__c + "'),',') as \"webVisitTotal\",\n      array_to_string(array(SELECT count(1) FROM sfdc.opportunity where recordtypeid = t1.recordtypeid and accountid = t1.accountid and NewsletterProperty__c = t1.accountid and VisitDate__c <= '" + this.theWeek.PeriodTo__c + "'),',') as \"newsVisitTotal\",\n      array_to_string(array(SELECT count(1) FROM sfdc.opportunity where recordtypeid = t1.recordtypeid and accountid = t1.accountid and ENewsletterProperty__c = t1.accountid and JoinBCDate__c < VisitDate__c and VisitDate__c <= '" + this.theWeek.PeriodTo__c + "'),',') as \"webJoinedTotal\",\n      array_to_string(array(SELECT count(1) FROM sfdc.opportunity where recordtypeid = t1.recordtypeid and accountid = t1.accountid and NewsletterProperty__c = t1.accountid and JoinBCDate__c < VisitDate__c and VisitDate__c <= '" + this.theWeek.PeriodTo__c + "'),',') as \"newsJoinedTotal\",\n      array_to_string(array(SELECT count(1) FROM sfdc.opportunity where recordtypeid = t1.recordtypeid and accountid = t1.accountid and ENewsletterProperty__c = t1.accountid and JoinBCDate__c = VisitDate__c and VisitDate__c <= '" + this.theWeek.PeriodTo__c + "'),',') as \"webLiveJoinTotal\",\n      array_to_string(array(SELECT count(1) FROM sfdc.opportunity where recordtypeid = t1.recordtypeid and accountid = t1.accountid and NewsletterProperty__c = t1.accountid and JoinBCDate__c = VisitDate__c and VisitDate__c <= '" + this.theWeek.PeriodTo__c + "'),',') as \"newsLiveJoinTotal\"\n    FROM sfdc.opportunity t1 where t1.accountid = '" + WeeklyRptConfig.projectId + "' and t1.recordtypeid = '" + BCRecordTypeId + "' limit 1")];
                    case 3:
                        bcRes = _d.sent();
                        BCTotalInfo = bcRes.rows[0];
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
    ExcelService.prototype.setSubQuestion = function (weeklyRptCfg) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var sebQuestionInfo, query, result, _a, _b, oppInfo, res;
            var e_22, _c;
            return tslib_1.__generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        if (!this.worksheet || !this.theWeek)
                            return [2];
                        sebQuestionInfo = {
                            weekPriceImpression: 0,
                            weekPriceImpressionCount: 0,
                            weekLocalImpression: 0,
                            weekLocalImpressionCount: 0,
                            weekExaminationDegree: 0,
                            weekExaminationDegreeCount: 0,
                            weekMREvaluation: 0,
                            weekMREvaluationCount: 0,
                            totalPriceImpression: 0,
                            totalLocalImpression: 0,
                            totalExaminationDegree: 0,
                            totalMREvaluation: 0,
                            PriceImpression: ["やや高い", "高い"],
                            ExaminationDegree: ["前向きに検討したい", "是非購入したい"],
                            LocalImpression: ["まあまあ良い", "大変良い"],
                            MREvaluation: ["まあまあ良い", "大変良い"]
                        };
                        query = "SELECT LocalImpression__c \"LocalImpression__c\",\n      MREvaluation__c \"MREvaluation__c\",\n      PriceImpression__c \"PriceImpression__c\",\n      ExaminationDegree__c \"ExaminationDegree__c\",\n      ResponseDate__c \"ResponseDate__c\"\n    FROM sfdc.opportunity\n    WHERE responsedate__c >= '" + this.theWeek.PeriodFrom__c + "' and responsedate__c <= '" + this.theWeek.PeriodTo__c + "' and accountid = '" + weeklyRptCfg.projectId + "' and (LocalImpression__c is not null or MREvaluation__c is not null or PriceImpression__c is not null or ExaminationDegree__c is not null)";
                        return [4, this.pgService.query(query)];
                    case 1:
                        result = _d.sent();
                        try {
                            for (_a = tslib_1.__values(result.rows), _b = _a.next(); !_b.done; _b = _a.next()) {
                                oppInfo = _b.value;
                                if (oppInfo.PriceImpression__c) {
                                    if (sebQuestionInfo.PriceImpression.includes(oppInfo.PriceImpression__c)) {
                                        sebQuestionInfo.weekPriceImpression++;
                                    }
                                    sebQuestionInfo.weekPriceImpressionCount++;
                                }
                                if (oppInfo.MREvaluation__c) {
                                    if (sebQuestionInfo.MREvaluation.includes(oppInfo.MREvaluation__c)) {
                                        sebQuestionInfo.weekMREvaluation++;
                                    }
                                    sebQuestionInfo.weekMREvaluationCount++;
                                }
                                if (oppInfo.LocalImpression__c) {
                                    if (sebQuestionInfo.LocalImpression.includes(oppInfo.LocalImpression__c)) {
                                        sebQuestionInfo.weekLocalImpression++;
                                    }
                                    sebQuestionInfo.weekLocalImpressionCount++;
                                }
                                if (oppInfo.ExaminationDegree__c) {
                                    if (sebQuestionInfo.ExaminationDegree.includes(oppInfo.ExaminationDegree__c)) {
                                        sebQuestionInfo.weekExaminationDegree++;
                                    }
                                    sebQuestionInfo.weekExaminationDegreeCount++;
                                }
                            }
                        }
                        catch (e_22_1) { e_22 = { error: e_22_1 }; }
                        finally {
                            try {
                                if (_b && !_b.done && (_c = _a.return)) _c.call(_a);
                            }
                            finally { if (e_22) throw e_22.error; }
                        }
                        this.worksheet.getCell("D53").value = sebQuestionInfo.weekPriceImpression;
                        this.worksheet.getCell("E53").value = sebQuestionInfo.weekPriceImpression / sebQuestionInfo.weekPriceImpressionCount;
                        this.worksheet.getCell("K53").value = sebQuestionInfo.weekLocalImpression;
                        this.worksheet.getCell("L53").value = sebQuestionInfo.weekLocalImpression / sebQuestionInfo.weekLocalImpressionCount;
                        this.worksheet.getCell("D54").value = sebQuestionInfo.weekExaminationDegree;
                        this.worksheet.getCell("E54").value =
                            sebQuestionInfo.weekExaminationDegree / sebQuestionInfo.weekExaminationDegreeCount;
                        this.worksheet.getCell("K54").value = sebQuestionInfo.weekMREvaluation;
                        this.worksheet.getCell("L54").value = sebQuestionInfo.weekMREvaluation / sebQuestionInfo.weekMREvaluationCount;
                        query = "SELECT\n        accountid as \"AccountId\",\n        array_to_string(array(SELECT count(1) FROM sfdc.opportunity where accountid = t1.accountid and responsedate__c <= '" + this.theWeek.PeriodTo__c + "' and LocalImpression__c in " + sfdc_service_1.SfdcService.getInSql(sebQuestionInfo.LocalImpression) + "),',') as \"LocalImpressionTotal\",\n        array_to_string(array(SELECT count(1) FROM sfdc.opportunity where accountid = t1.accountid and responsedate__c <= '" + this.theWeek.PeriodTo__c + "' and LocalImpression__c is not null),',') as \"LocalImpressionAll\",\n        array_to_string(array(SELECT count(1) FROM sfdc.opportunity where accountid = t1.accountid and responsedate__c <= '" + this.theWeek.PeriodTo__c + "' and MREvaluation__c in " + sfdc_service_1.SfdcService.getInSql(sebQuestionInfo.MREvaluation) + "),',') as \"MREvaluationTotal\",\n        array_to_string(array(SELECT count(1) FROM sfdc.opportunity where accountid = t1.accountid and responsedate__c <= '" + this.theWeek.PeriodTo__c + "' and MREvaluation__c is not null),',') as \"MREvaluationAll\",\n        array_to_string(array(SELECT count(1) FROM sfdc.opportunity where accountid = t1.accountid and responsedate__c <= '" + this.theWeek.PeriodTo__c + "' and PriceImpression__c in " + sfdc_service_1.SfdcService.getInSql(sebQuestionInfo.PriceImpression) + "),',') as \"PriceImpressionTotal\",\n        array_to_string(array(SELECT count(1) FROM sfdc.opportunity where accountid = t1.accountid and responsedate__c <= '" + this.theWeek.PeriodTo__c + "' and PriceImpression__c is not null),',') as \"PriceImpressionAll\",\n        array_to_string(array(SELECT count(1) FROM sfdc.opportunity where accountid = t1.accountid and responsedate__c <= '" + this.theWeek.PeriodTo__c + "' and ExaminationDegree__c in " + sfdc_service_1.SfdcService.getInSql(sebQuestionInfo.ExaminationDegree) + "),',') as \"ExaminationDegreeTotal\",\n        array_to_string(array(SELECT count(1) FROM sfdc.opportunity where accountid = t1.accountid and responsedate__c <= '" + this.theWeek.PeriodTo__c + "' and ExaminationDegree__c is not null),',') as \"ExaminationDegreeAll\"\n    FROM sfdc.opportunity t1 where t1.accountid = '0011m00000SvSEPAA3' limit 1";
                        return [4, this.pgService.query(query)];
                    case 2:
                        result = _d.sent();
                        if (result.rowCount === 0)
                            return [2];
                        res = result.rows[0];
                        this.worksheet.getCell("F53").value = Number(res.PriceImpressionTotal);
                        this.worksheet.getCell("G53").value = Number(res.PriceImpressionTotal) / Number(res.PriceImpressionAll);
                        this.worksheet.getCell("M53").value = Number(res.LocalImpressionTotal);
                        this.worksheet.getCell("N53").value = Number(res.LocalImpressionTotal) / Number(res.LocalImpressionAll);
                        this.worksheet.getCell("F54").value = Number(res.ExaminationDegreeTotal);
                        this.worksheet.getCell("G54").value = Number(res.ExaminationDegreeTotal) / Number(res.ExaminationDegreeAll);
                        this.worksheet.getCell("M54").value = Number(res.MREvaluationTotal);
                        this.worksheet.getCell("N54").value = Number(res.MREvaluationTotal) / Number(res.MREvaluationAll);
                        return [2];
                }
            });
        });
    };
    ExcelService.prototype.setOOSInfo = function (weeklyRptCfg) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var query, result, res;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.worksheet || !this.theWeek)
                            return [2];
                        query = "SELECT\n      accountid as \"AccountId\",\n      array_to_string(array(SELECT count(1) FROM sfdc.opportunity where accountid = t1.accountid and responsedate__c <= visitreservedate__c and StageName = 'OSS(\u5951\u7D04\u524D)'),',') as \"visitreserve\",\n      array_to_string(array(SELECT count(1) FROM sfdc.opportunity where accountid = t1.accountid and responsedate__c <= visitreservedate__c and StageName = 'OSS(\u5951\u7D04\u524D)' and VisitDate__c is not null),',') as \"visit\"\n    FROM sfdc.opportunity t1 where t1.accountid = '" + weeklyRptCfg.projectId + "' limit 1";
                        return [4, this.pgService.query(query)];
                    case 1:
                        result = _a.sent();
                        if (result.rowCount === 0)
                            return [2];
                        res = result.rows[0];
                        this.worksheet.getCell("B57").value = Number(res.visitreserve);
                        this.worksheet.getCell("D57").value = Number(res.visit);
                        this.worksheet.getCell("F57").value = res.visit / res.visitreserve;
                        return [2];
                }
            });
        });
    };
    ExcelService.prototype.setWeekTableInfo = function (sfdcObj, weekDateMap, filedName, weekDateFiledName) {
        if (!sfdcObj[filedName])
            return;
        var dateKey = sfdcObj[filedName].substring(0, 10);
        var weekTable = weekDateMap.get(dateKey);
        if (weekTable) {
            weekTable[weekDateFiledName ? weekDateFiledName : filedName]++;
        }
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
            if (filedName === "InquiryDateFirst__c" && sfdcObj.VisitDate__c && weekTable) {
                if (weekTable.InquiryVisitNum) {
                    weekTable.InquiryVisitNum++;
                }
                else {
                    weekTable.InquiryVisitNum = 1;
                }
            }
        }
        if (weekTable)
            weekTable[weekDateFiledName ? weekDateFiledName : filedName]++;
    };
    return ExcelService;
}());
exports.ExcelService = ExcelService;
//# sourceMappingURL=excel.service.js.map