"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.BcCommonController = void 0;
var tslib_1 = require("tslib");
var express_1 = require("express");
var sfdc_service_1 = require("../service/sfdc.service");
var BcCommonController = (function () {
    function BcCommonController() {
        var _this = this;
        this.sfdcService = new sfdc_service_1.SfdcService();
        this.initBcJoin = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var hopeAreaObj, hopeRoomList, budgetPriceList, hopeSquareList, bulletinMediaList, mailFormatList, hopeDeliveryList, viewModelList, genderList, overseaList, pickListMulti, _a, _b, quesObj, quesNum, optionList, _c, _d, opObj, option, resultJson;
            var e_1, _e, e_2, _f;
            return tslib_1.__generator(this, function (_g) {
                switch (_g.label) {
                    case 0:
                        hopeAreaObj = {};
                        hopeRoomList = [];
                        budgetPriceList = [];
                        hopeSquareList = [];
                        bulletinMediaList = [];
                        mailFormatList = [];
                        hopeDeliveryList = [];
                        viewModelList = [];
                        genderList = [];
                        overseaList = [];
                        return [4, this.getBcPickMultiList("BC入会", "'A0005','C0111','B0009','B0119','B0120','C0112','B0166','B0168','B0167','B0005'")];
                    case 1:
                        pickListMulti = _g.sent();
                        if (pickListMulti.records.length > 0) {
                            try {
                                for (_a = tslib_1.__values(pickListMulti.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                    quesObj = _b.value;
                                    quesNum = quesObj.QuestionNumber__c;
                                    if (quesNum == "C0111") {
                                        hopeAreaObj = quesObj;
                                    }
                                    else {
                                        optionList = quesObj.optionProjectQuestionnaireQuestions__r;
                                        if (optionList && optionList.records.length > 0) {
                                            try {
                                                for (_c = (e_2 = void 0, tslib_1.__values(optionList.records)), _d = _c.next(); !_d.done; _d = _c.next()) {
                                                    opObj = _d.value;
                                                    option = { value: opObj.OptionItemNumber__c, label: opObj.OptionItemValue__c };
                                                    if (quesObj.QuestionNumber__c == "A0005") {
                                                        overseaList.push(option);
                                                    }
                                                    if (quesObj.QuestionNumber__c == "B0005") {
                                                        genderList.push(option);
                                                    }
                                                    if (quesObj.QuestionNumber__c == "B0167") {
                                                        viewModelList.push(option);
                                                    }
                                                    if (quesObj.QuestionNumber__c == "B0009") {
                                                        hopeRoomList.push(option);
                                                    }
                                                    if (quesObj.QuestionNumber__c == "B0119") {
                                                        option = { value: opObj.OptionItemNumber__c, label: opObj.Label__c };
                                                        budgetPriceList.push(option);
                                                    }
                                                    if (quesObj.QuestionNumber__c == "B0120") {
                                                        option = { value: opObj.OptionItemNumber__c, label: opObj.Label__c };
                                                        hopeSquareList.push(option);
                                                    }
                                                    if (quesObj.QuestionNumber__c == "C0112") {
                                                        bulletinMediaList.push(option);
                                                    }
                                                    if (quesObj.QuestionNumber__c == "B0166") {
                                                        mailFormatList.push(option);
                                                    }
                                                    if (quesObj.QuestionNumber__c == "B0168") {
                                                        hopeDeliveryList.push(option);
                                                    }
                                                }
                                            }
                                            catch (e_2_1) { e_2 = { error: e_2_1 }; }
                                            finally {
                                                try {
                                                    if (_d && !_d.done && (_f = _c.return)) _f.call(_c);
                                                }
                                                finally { if (e_2) throw e_2.error; }
                                            }
                                        }
                                    }
                                }
                            }
                            catch (e_1_1) { e_1 = { error: e_1_1 }; }
                            finally {
                                try {
                                    if (_b && !_b.done && (_e = _a.return)) _e.call(_a);
                                }
                                finally { if (e_1) throw e_1.error; }
                            }
                        }
                        resultJson = {
                            budgetPriceList: budgetPriceList,
                            hopeRoomList: hopeRoomList,
                            hopeAreaObj: hopeAreaObj,
                            hopeSquareList: hopeSquareList,
                            bulletinMediaList: bulletinMediaList,
                            mailFormatList: mailFormatList,
                            hopeDeliveryList: hopeDeliveryList,
                            viewModelList: viewModelList,
                            genderList: genderList,
                            overseaList: overseaList
                        };
                        res.json(resultJson);
                        return [2];
                }
            });
        }); };
        this.sendBcJoin = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var state, type, subjectDetail, quesResult, contentDetail, bcCreateFlg, projectNumber, fromMail, campaignCode, detail, nowDate, room, dateTime, jsonData;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        state = req.body.params.state;
                        type = req.body.params.type;
                        subjectDetail = "BC入会";
                        if (type == "event") {
                            subjectDetail = "BC入会(イベント)";
                        }
                        return [4, this.getProjectQuesForBc(subjectDetail)];
                    case 1:
                        quesResult = _a.sent();
                        contentDetail = {};
                        bcCreateFlg = false;
                        projectNumber = "";
                        fromMail = "";
                        campaignCode = "";
                        if (quesResult.totalSize > 0) {
                            projectNumber = quesResult.records[0].ProjectNumber__c;
                            fromMail = quesResult.records[0].Project__r.PropertyMail__c;
                            detail = "{";
                            detail += '"A0001":"' + state.memberSeiKanji + '",';
                            detail += '"A0002":"' + state.memberMeiKanji + '",';
                            detail += '"A0003":"' + state.memberSeiKana + '",';
                            detail += '"A0004":"' + state.memberMeiKana + '",';
                            detail += '"A0007":"' + state.prefecture + '",';
                            detail += '"A0014":"' + state.memberMail.value + '",';
                            detail += '"A0016":"' + state.memberBirthdayYear + "-" + state.memberBirthdayMonth + "-" + state.memberBirthdayDay + '",';
                            nowDate = this.getSystemDate("");
                            if (type == "bc") {
                                campaignCode = state.campaignCode.value;
                                detail += "\"A0005\":{\"" + this.getOptionKey(state.abroadType.value, state.abroadType.pickList) + "\":\"\"},";
                                detail += '"A0006":"' + state.zipCode1 + state.zipCode2 + '",';
                                detail += '"A0008":"' + state.city + '",';
                                detail += '"A0009":"' + state.towns + '",';
                                detail += '"A0010":"' + state.address + '",';
                                room = state.roomNumPropertyName;
                                if (state.abroadType.value == "海外") {
                                    room = state.overseaAddress;
                                }
                                detail += '"A0011":"' + room + '",';
                                detail += '"B0166":{"' + state.mailFormat.value + '":""},';
                                detail += "\"B0167\":{\"" + this.getOptionKey(state.viewType, state.viewModelList) + "\":\"\"},";
                                detail += "\"B0005\":{\"" + this.getOptionKey(state.memberGender.value, state.memberGender.pickList) + "\":\"\"},";
                                detail += '"A0013":"' + state.phoneNumber.value + '",';
                                detail += "\"B0009\":{\"" + this.getOptionKey(state.hopeRoomType.value, state.hopeRoomType.pickList) + "\":\"\"},";
                                detail += "\"B0119\":{\"" + this.getOptionKey(state.budgetPrice.value, state.budgetPrice.pickList) + "\":\"\"},";
                                detail += "\"B0120\":{\"" + this.getOptionKey(state.hopeSquare.value, state.hopeSquare.pickList) + "\":\"\"},";
                                detail += "\"B0166\":{\"" + this.getOptionKey(state.mailFormat.value, state.mailFormat.pickList) + "\":\"\"},";
                                detail += "\"B0168\":{\"" + this.getOptionKey(state.hopeDelivery.value, state.hopeDelivery.pickList) + "\":\"\"},";
                                if (state.hopeDelivery.value.indexOf("希望する") != -1) {
                                    detail += "\"B0201\":\"" + nowDate + "\",";
                                    detail += "\"B0202\":\"BC\u5165\u4F1A\u30D5\u30A9\u30FC\u30E0\",";
                                }
                            }
                            detail += "\"B0207\":\"BC\u4F1A\u54E1\",";
                            detail += "\"B0195\":\"" + nowDate + "\",";
                            if (type == "bc") {
                                detail += "\"B0196\":\"BC\u5165\u4F1A\u30D5\u30A9\u30FC\u30E0\",";
                            }
                            else {
                                detail += "\"B0196\":\"\u30A4\u30D9\u30F3\u30C8\",";
                            }
                            detail = detail.substring(0, detail.length - 1);
                            detail += "}";
                            contentDetail = JSON.parse(detail);
                        }
                        dateTime = this.getSystemDate("time");
                        jsonData = {
                            ProjectNumber: projectNumber,
                            AnswerCode: "",
                            CampaignCode: campaignCode,
                            CreateBC: bcCreateFlg,
                            AnswerDate: dateTime,
                            Create: dateTime,
                            BOCID: "",
                            Content: contentDetail
                        };
                        console.log(JSON.stringify(jsonData));
                        this.sendMail(fromMail, state.toMail, "フォーム投稿", jsonData);
                        res.send("ok");
                        return [2];
                }
            });
        }); };
        this.initWithdrawBc = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var withdrawResonList, overseaList, pickList, _a, _b, quesObj, optionList, _c, _d, opObj, option, resultJson;
            var e_3, _e, e_4, _f;
            return tslib_1.__generator(this, function (_g) {
                switch (_g.label) {
                    case 0:
                        withdrawResonList = [];
                        overseaList = [];
                        return [4, this.getBcPickList("BC退会", "'B0176','A0005'")];
                    case 1:
                        pickList = _g.sent();
                        if (pickList.records.length > 0) {
                            try {
                                for (_a = tslib_1.__values(pickList.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                    quesObj = _b.value;
                                    optionList = quesObj.optionProjectQuestionnaireQuestions__r;
                                    if (optionList && optionList.records.length > 0) {
                                        try {
                                            for (_c = (e_4 = void 0, tslib_1.__values(optionList.records)), _d = _c.next(); !_d.done; _d = _c.next()) {
                                                opObj = _d.value;
                                                option = { value: opObj.OptionItemNumber__c, label: opObj.OptionItemValue__c };
                                                if (quesObj.QuestionNumber__c == "B0176") {
                                                    withdrawResonList.push(option);
                                                }
                                                if (quesObj.QuestionNumber__c == "A0005") {
                                                    overseaList.push(option);
                                                }
                                            }
                                        }
                                        catch (e_4_1) { e_4 = { error: e_4_1 }; }
                                        finally {
                                            try {
                                                if (_d && !_d.done && (_f = _c.return)) _f.call(_c);
                                            }
                                            finally { if (e_4) throw e_4.error; }
                                        }
                                    }
                                }
                            }
                            catch (e_3_1) { e_3 = { error: e_3_1 }; }
                            finally {
                                try {
                                    if (_b && !_b.done && (_e = _a.return)) _e.call(_a);
                                }
                                finally { if (e_3) throw e_3.error; }
                            }
                        }
                        resultJson = {
                            withdrawResonList: withdrawResonList,
                            overseaList: overseaList
                        };
                        res.json(resultJson);
                        return [2];
                }
            });
        }); };
        this.snedWithdrawBc = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var state, quesResult, contentDetail, bcCreateFlg, projectNumber, fromMail, selectOpp, oppResult, detail, withdrawDate, oppObj, dateTime, jsonData;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        state = req.body.params.state;
                        return [4, this.getProjectQuesForBc("BC退会")];
                    case 1:
                        quesResult = _a.sent();
                        contentDetail = {};
                        bcCreateFlg = false;
                        projectNumber = "";
                        fromMail = "";
                        selectOpp = "select Id,Opportunity__r.ENewsletterAppDate__c,Opportunity__r.NewsletterAppDate__c\n    from ContactEmailList__c where MailAddress__c = '" + state.memberMail.value + "'\n    and Opportunity__r.Contact__r.LastName='" + state.memberSeiKanji + "'\n    and Opportunity__r.Contact__r.FirstName='" + state.memberMeiKanji + "'\n    and Opportunity__r.Contact__r.FamilyNameKana__c='" + state.memberSeiKana + "'\n    and Opportunity__r.Contact__r.FirstNameKana__c='" + state.memberMeiKana + "'\n    and AccountType__c = 'BC\u7528'";
                        return [4, this.sfdcService.query(selectOpp)];
                    case 2:
                        oppResult = _a.sent();
                        if (oppResult.totalSize > 0) {
                            if (quesResult.totalSize > 0) {
                                projectNumber = quesResult.records[0].ProjectNumber__c;
                                bcCreateFlg = quesResult.records[0].BcMembershipCreation__c;
                                fromMail = quesResult.records[0].Project__r.PropertyMail__c;
                                detail = "{";
                                detail += '"A0001":"' + state.memberSeiKanji + '",';
                                detail += '"A0002":"' + state.memberMeiKanji + '",';
                                detail += '"A0003":"' + state.memberSeiKana + '",';
                                detail += '"A0004":"' + state.memberMeiKana + '",';
                                detail += '"A0014":"' + state.memberMail.value + '",';
                                detail += "\"A0005\":{\"" + this.getOptionKey(state.abroadType.value, state.abroadType.pickList) + "\":\"\"},";
                                detail += '"A0006":"' + state.zipCode1 + state.zipCode2 + '",';
                                detail += '"A0007":"' + state.prefecture + '",';
                                detail += '"A0008":"' + state.city + '",';
                                detail += '"A0009":"' + state.towns + '",';
                                detail += '"A0010":"' + state.address + '",';
                                detail += '"A0011":"' + state.roomNumPropertyName + '",';
                                detail += "\"B0176\":{\"" + this.getOptionKey(state.withdrawReson.value, state.withdrawReson.pickList) + "\":\"\"},";
                                withdrawDate = this.getSystemDate("").replaceAll("/", "-");
                                if (oppResult.totalSize > 0) {
                                    oppObj = oppResult.records[0];
                                    if (oppObj.Opportunity__r.ENewsletterAppDate__c) {
                                        detail += '"B0200":"' + withdrawDate + '",';
                                    }
                                    if (oppObj.Opportunity__r.NewsletterAppDate__c) {
                                        detail += '"B0206":"' + withdrawDate + '",';
                                    }
                                }
                                detail += '"B0175":"' + withdrawDate + '"';
                                detail += "}";
                                console.log(detail);
                                contentDetail = JSON.parse(detail);
                                dateTime = this.getSystemDate("time");
                                jsonData = {
                                    ProjectNumber: projectNumber,
                                    AnswerCode: "",
                                    CreateBC: bcCreateFlg,
                                    AnswerDate: dateTime,
                                    Create: dateTime,
                                    BOCID: state.memberMyPageId,
                                    Content: contentDetail
                                };
                                console.log(JSON.stringify(jsonData));
                                this.sendMail(fromMail, state.toMail, "フォーム投稿", jsonData);
                                res.send("ok");
                            }
                        }
                        else {
                            res.send("ng");
                        }
                        return [2];
                }
            });
        }); };
        this.getProjectQuesForBc = function (subjectDetail) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var selectQues, quesResult;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        selectQues = "select ProjectNumber__c ,BcMembershipCreation__c,Project__r.PropertyMail__c,\n    (select QuestionNumber__c,QuestionContent__c from Questionnaire__r) from ProjectQuestionnaire__c\n     where OppRecordType__c ='BC' and Project__r.ProjectType__c='BC\u7528'\n     and ProjectType__c = '" + subjectDetail + "'";
                        return [4, this.sfdcService.query(selectQues)];
                    case 1:
                        quesResult = _a.sent();
                        return [2, quesResult];
                }
            });
        }); };
        this.getBcPickList = function (projectType, qesNumList) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var selectUser, quesResult;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        selectUser = "select Id ,QuestionNumber__c,\n      (select Id, OptionItemValue__c,OptionItemNumber__c from optionProjectQuestionnaireQuestions__r\n        where OptionItemDeleteFlag__c=false\n        order by MainItemOrderNumber__c,OptionItemOrderNumber__c )\n      from ProjectQuestionnaireQuestions__c\n      where QuestionnaireAccount__r.ProjectType__c ='" + projectType + "' and QuestionDeleteFlag__c = false\n      and QuestionnaireAccount__r.Project__r.ProjectType__c='BC\u7528'\n      and QuestionNumber__c in (" + qesNumList + ") order by OrderNumber__c ";
                        return [4, this.sfdcService.query(selectUser)];
                    case 1:
                        quesResult = _a.sent();
                        return [2, quesResult];
                }
            });
        }); };
        this.getBcPickMultiList = function (projectType, qesNumList) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var selectUser, quesResult;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        selectUser = "select Id, QuestionNumber__c,MaxRequiredChoicesNumber__c,\n      (select Id,QuestionnaireMainMaster__c,QuestionnaireItemMaster__c,MainItemValue__c,\n        ProjectQuestionnaireQuestions__c,OptionItemValue__c,OptionItemNumber__c,Label__c,\n        OptionItemOrderNumber__c\n        from optionProjectQuestionnaireQuestions__r where MainItemDeleteFlag__c=false and OptionItemDeleteFlag__c=false\n         order by MainItemOrderNumber__c,OptionItemOrderNumber__c )\n      from ProjectQuestionnaireQuestions__c\n      where QuestionnaireAccount__r.ProjectType__c ='" + projectType + "' and QuestionDeleteFlag__c = false\n      and QuestionnaireAccount__r.Project__r.ProjectType__c='BC\u7528'\n      and QuestionNumber__c in (" + qesNumList + ") order by OrderNumber__c ";
                        return [4, this.sfdcService.query(selectUser)];
                    case 1:
                        quesResult = _a.sent();
                        return [2, quesResult];
                }
            });
        }); };
        this.router = express_1.Router();
        this.intializeRoutes();
    }
    BcCommonController.prototype.intializeRoutes = function () {
        this.router.post("/initBcJoin", this.initBcJoin);
        this.router.post("/sendBcJoin", this.sendBcJoin);
        this.router.post("/initWithdrawBc", this.initWithdrawBc);
        this.router.post("/snedWithdrawBc", this.snedWithdrawBc);
    };
    BcCommonController.prototype.sendMail = function (fromMail, toMail, subjectDetail, jsonData) {
        if (fromMail && toMail) {
            var sendmail = require("sendmail")();
            sendmail({
                from: fromMail,
                to: toMail,
                subject: subjectDetail,
                text: JSON.stringify(jsonData)
            }, function (err, reply) {
                if (err) {
                    console.log("send mail fail");
                }
            });
        }
    };
    BcCommonController.prototype.getSystemDate = function (type) {
        var nowtime = new Date();
        var month = nowtime.getMonth() + 1 + "";
        if (nowtime.getMonth() + 1 < 10) {
            month = "0" + month;
        }
        var days = nowtime.getDate() + "";
        if (nowtime.getDate() < 10) {
            days = "0" + days;
        }
        var hours = nowtime.getHours() + "";
        if (nowtime.getHours() < 10) {
            hours = "0" + hours;
        }
        var minutes = nowtime.getMinutes() + "";
        if (nowtime.getMinutes() < 10) {
            minutes = "0" + minutes;
        }
        var seconds = nowtime.getSeconds() + "";
        if (nowtime.getSeconds() < 10) {
            seconds = "0" + seconds;
        }
        var dateTime = nowtime.getFullYear() + "-" + month + "-" + days;
        if (type == "time") {
            dateTime += " " + hours + ":" + minutes + ":" + seconds;
        }
        return dateTime;
    };
    BcCommonController.prototype.nullToString = function (target) {
        var resStr = "";
        if (target != undefined && target != null && target != "") {
            resStr = target;
        }
        return resStr;
    };
    BcCommonController.prototype.getOptionKey = function (label, targetList) {
        var e_5, _a;
        var resVal = label;
        try {
            for (var targetList_1 = tslib_1.__values(targetList), targetList_1_1 = targetList_1.next(); !targetList_1_1.done; targetList_1_1 = targetList_1.next()) {
                var opObj = targetList_1_1.value;
                if (label == opObj.label) {
                    resVal = opObj.value;
                }
            }
        }
        catch (e_5_1) { e_5 = { error: e_5_1 }; }
        finally {
            try {
                if (targetList_1_1 && !targetList_1_1.done && (_a = targetList_1.return)) _a.call(targetList_1);
            }
            finally { if (e_5) throw e_5.error; }
        }
        return resVal;
    };
    return BcCommonController;
}());
exports.BcCommonController = BcCommonController;
//# sourceMappingURL=bc.common.controller.js.map