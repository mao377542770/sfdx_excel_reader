"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.BocQuestionaryController = void 0;
var tslib_1 = require("tslib");
var express_1 = require("express");
var sfdc_service_1 = require("../service/sfdc.service");
var BocQuestionaryController = (function () {
    function BocQuestionaryController() {
        var _this = this;
        this.sfdcService = new sfdc_service_1.SfdcService();
        this.projectLogin = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var state, projectCode, password, selectQues, quesResult, resJson;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        state = req.body.params.state;
                        projectCode = state.projectCode;
                        password = state.password;
                        selectQues = "select Id, ProjectCode__c,Name from Account " + " where ProjectCode__c = '" + projectCode + "'";
                        if (password != "" && password != null) {
                            selectQues += " and ProjectPassword__c = '" + password + "'";
                        }
                        return [4, this.sfdcService.query(selectQues)];
                    case 1:
                        quesResult = _a.sent();
                        resJson = {
                            result: "NG",
                            proName: ""
                        };
                        if (quesResult.totalSize > 0) {
                            resJson.result = "OK";
                            resJson.proName = quesResult.records[0].Name;
                        }
                        res.send(resJson);
                        return [2];
                }
            });
        }); };
        this.getCustomerInfo = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var custPid, projectCode, selectQuery, selectRresult, selectBc, bcRresult, resultJson;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        custPid = req.body.params.custPid;
                        projectCode = req.body.params.projectCode;
                        selectQuery = "SELECT Id,ContactName__c,ContactNameKana__c,Contact__r.LastName,Contact__r.FirstName,\n      Contact__r.FamilyNameKana__c,Contact__r.FirstNameKana__c,MemberMyPageID__c,VisitDate__c,\n      Contact__r.Gender__c,Contact__r.IndividualCorporate__c,\n      CaseNo__c,VisitNo__c,ContactBirthday__c,Prefectures__c,City__c,Towns__c,Address__c,RoomNumPropertyName__c,\n      (select Id, MailAddress__c,ContactPhoneNumber__c from contactEmailOpportunity__r where MainFlag__c = true)\n       FROM Opportunity WHERE ContactPid__c = '" + custPid + "' and Account.ProjectCode__c= '" + projectCode + "'";
                        return [4, this.sfdcService.query(selectQuery)];
                    case 1:
                        selectRresult = _a.sent();
                        selectBc = "select Id,ENewsletterAppDate__c,NewsletterAppDate__c\n     from Opportunity WHERE ContactPid__c = '" + custPid + "' and RecordType.DeveloperName='BC'\n    and RecordType.SobjectType='Opportunity'";
                        return [4, this.sfdcService.query(selectBc)];
                    case 2:
                        bcRresult = _a.sent();
                        resultJson = {
                            customerInfo: selectRresult,
                            bcMember: bcRresult.records
                        };
                        res.json(resultJson);
                        return [2];
                }
            });
        }); };
        this.searchStation = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var stationNameForSearch, selectQuery, selectRresult;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        stationNameForSearch = req.body.params.stationNameForSearch;
                        selectQuery = "SELECT Id,StationName__c,StationLineName__c FROM EkiMst__c WHERE StationName__c like '%" + stationNameForSearch + "%' ";
                        return [4, this.sfdcService.query(selectQuery)];
                    case 1:
                        selectRresult = _a.sent();
                        res.send(selectRresult);
                        return [2];
                }
            });
        }); };
        this.getProjectQuestions = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var projectCode, quesType, selectUser, quesResult, loginRes, typeList, searchPickFlg, _a, _b, quesObj, answerType, searchPick, projectQuery, pickListResult, selectTeam, teamResult;
            var e_1, _c;
            return tslib_1.__generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        projectCode = req.body.projectCode;
                        quesType = req.body.quesType;
                        selectUser = "select Id, AnswerType__c ,QuestionnaireQuestionMaster__r.Required__c,AnswerTypeRestrictionsText__c,\n      MaxRequiredChoicesNumber__c,QuestionnaireAccount__r.ProjectNumber__c,QuestionnaireAccount__r.Project__r.PropertyMail__c,\n      QuestionContent__c,QuestionNumber__c,OrderNumber__c,QuestionnaireQuestionMaster__c,QuestionnaireQuestionMaster__r.Unit__c,\n      QuestionnaireAccount__r.BcMembershipCreation__c,MinRequiredChoicesNumber__c,QuestionnaireQuestionMaster__r.Name,\n      QuestionnaireAccount__r.Project__r.Name,QuestionnaireAccount__r.Project__c,\n      (select Id,QuestionnaireMainMaster__c,QuestionnaireItemMaster__c,MainItemValue__c,\n        ProjectQuestionnaireQuestions__c,OptionItemValue__c,OptionItemNumber__c,\n        OptionItemOrderNumber__c,QuestionnaireItemMaster__r.NeedText__c\n        from optionProjectQuestionnaireQuestions__r where MainItemDeleteFlag__c=false and OptionItemDeleteFlag__c=false\n         order by MainItemOrderNumber__c,OptionItemOrderNumber__c )\n      from ProjectQuestionnaireQuestions__c\n      where QuestionnaireAccount__r.ProjectType__c ='" + quesType + "' and QuestionDeleteFlag__c=false\n      and QuestionnaireAccount__r.Project__r.ProjectCode__c = '" + projectCode + "' order by OrderNumber__c ";
                        return [4, this.sfdcService.query(selectUser)];
                    case 1:
                        quesResult = _d.sent();
                        loginRes = {
                            result: quesResult.records,
                            pickList: [{}],
                            teamList: [{}]
                        };
                        if (!(quesResult.totalSize > 0)) return [3, 3];
                        typeList = [];
                        searchPickFlg = false;
                        try {
                            for (_a = tslib_1.__values(quesResult.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                quesObj = _b.value;
                                answerType = quesObj.AnswerType__c;
                                if (answerType == "選択リスト" || answerType == "複数選択リスト") {
                                    searchPickFlg = true;
                                    typeList.push("'" + quesObj.QuestionnaireQuestionMaster__c + "'");
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
                        if (!searchPickFlg) return [3, 3];
                        searchPick = "(" + typeList + ")";
                        projectQuery = "select Id,QuestionnaireMainMaster__c,QuestionnaireItemMaster__c,\n         ProjectQuestionnaireQuestions__c,OptionItemValue__c,OptionItemNumber__c,\n         OptionItemOrderNumber__c\n         from ProjectQuestionnaireQuestionsOption__c\n         where ProjectQuestionnaireQuestions__c in " + searchPick;
                        return [4, this.sfdcService.query(projectQuery)];
                    case 2:
                        pickListResult = _d.sent();
                        loginRes.pickList = pickListResult.records;
                        _d.label = 3;
                    case 3:
                        selectTeam = "select Id, AccountTeamMaster__c, AccountTeamMaster__r.LastName__c, AccountTeamMaster__r.FirstName__c,\n    AccountTeamMaster__r.Name\n    from ProjectAccountTeam__c where Project__r.ProjectCode__c = '" + projectCode + "'";
                        return [4, this.sfdcService.query(selectTeam)];
                    case 4:
                        teamResult = _d.sent();
                        console.log(teamResult.records);
                        loginRes.teamList = teamResult.records;
                        res.send(JSON.stringify(loginRes));
                        return [2];
                }
            });
        }); };
        this.sendQaAndCreatePdf = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var fromMail, toMail, dateTime, jsonData;
            return tslib_1.__generator(this, function (_a) {
                fromMail = req.body.params.projectMail;
                toMail = req.body.params.toMail;
                dateTime = this.getSystemDate();
                jsonData = req.body.params.jsonData;
                jsonData.AnswerDate = dateTime;
                jsonData.Create = dateTime;
                console.log(JSON.stringify(jsonData));
                this.sendMail(fromMail, toMail, "フォーム投稿", jsonData);
                res.send("OK");
                return [2];
            });
        }); };
        this.router = express_1.Router();
        this.intializeRoutes();
    }
    BocQuestionaryController.prototype.intializeRoutes = function () {
        this.router.post("/projectLogin", this.projectLogin);
        this.router.post("/getProjectQuestions", this.getProjectQuestions);
        this.router.post("/getCustomerInfo", this.getCustomerInfo);
        this.router.post("/searchStation", this.searchStation);
        this.router.post("/sendQaAndCreatePdf", this.sendQaAndCreatePdf);
    };
    BocQuestionaryController.prototype.sendMail = function (fromMail, toMail, subjectDetail, jsonData) {
        if (fromMail && toMail) {
            var sendmail = require("sendmail")();
            sendmail({
                from: fromMail,
                to: toMail,
                subject: subjectDetail,
                text: JSON.stringify(jsonData)
            }, function (err, reply) {
                console.log("send mail fail");
            });
        }
    };
    BocQuestionaryController.prototype.getSystemDate = function () {
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
        var dateTime = nowtime.getFullYear() + "-" + month + "-" + days + " " + hours + ":" + minutes + ":" + seconds;
        return dateTime;
    };
    return BocQuestionaryController;
}());
exports.BocQuestionaryController = BocQuestionaryController;
//# sourceMappingURL=boc.questionary.controller.js.map