"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.BocCommonController = void 0;
var tslib_1 = require("tslib");
var express_1 = require("express");
var sfdc_service_1 = require("../service/sfdc.service");
var BocCommonController = (function () {
    function BocCommonController() {
        var _this = this;
        this.sfdcService = new sfdc_service_1.SfdcService();
        this.initInquiry = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var memberMyPageId, type, selectQuery, selectRresult, newProjectList, yearTypeList, presentInfo, pickList, _a, _b, quesObj, optionList, _c, _d, opObj, option, pickList, _e, _f, quesObj, optionList, _g, _h, opObj, option, projectNumber, presentQuery, presentRes, resultJson, error_1;
            var e_1, _j, e_2, _k, e_3, _l, e_4, _m;
            return tslib_1.__generator(this, function (_o) {
                switch (_o.label) {
                    case 0:
                        _o.trys.push([0, 8, , 9]);
                        memberMyPageId = req.body.params.memberMyPageId;
                        type = req.body.params.type;
                        selectQuery = "SELECT Id,AccountId,Account.ProjectCode__c,Contact__r.LastName,Contact__r.FirstName, " +
                            "Contact__r.Birthdate,ContactName__c,ContactNameKana__c,Contact__r.FamilyNameKana__c,Contact__r.FirstNameKana__c," +
                            "ZipCode__c,Prefectures__c,Towns__c,RoomNumPropertyName__c,City__c,Address__c,RequestMailMemberNewsletter__c" +
                            ",AddressRelationship__c" +
                            ",(select MailAddress__c,ContactPhoneNumber__c from contactEmailOpportunity__r where MainFlag__c = true)" +
                            ",(select Project__c,BocProjectName__c,BocRoom__c from BOC__r )" +
                            " FROM Opportunity WHERE MemberMyPageID__c = '" + memberMyPageId + "'" +
                            " and MemberWithdrawalDate__c = null and RecordType.DeveloperName='BOC' and RecordType.SobjectType='Opportunity'";
                        return [4, this.sfdcService.query(selectQuery)];
                    case 1:
                        selectRresult = _o.sent();
                        newProjectList = [];
                        yearTypeList = [];
                        presentInfo = null;
                        if (!(type == "boc_buy")) return [3, 3];
                        return [4, this.getQuesPickList("特典：東京建物分譲物件ご紹介特典(新）", "'B0194'")];
                    case 2:
                        pickList = _o.sent();
                        if (pickList.records.length > 0) {
                            try {
                                for (_a = tslib_1.__values(pickList.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                    quesObj = _b.value;
                                    optionList = quesObj.optionProjectQuestionnaireQuestions__r;
                                    if (optionList && optionList.records.length > 0) {
                                        try {
                                            for (_c = (e_2 = void 0, tslib_1.__values(optionList.records)), _d = _c.next(); !_d.done; _d = _c.next()) {
                                                opObj = _d.value;
                                                option = { value: opObj.OptionItemNumber__c, label: opObj.OptionItemValue__c };
                                                if (quesObj.QuestionNumber__c == "B0194") {
                                                    newProjectList.push(option);
                                                }
                                            }
                                        }
                                        catch (e_2_1) { e_2 = { error: e_2_1 }; }
                                        finally {
                                            try {
                                                if (_d && !_d.done && (_k = _c.return)) _k.call(_c);
                                            }
                                            finally { if (e_2) throw e_2.error; }
                                        }
                                    }
                                }
                            }
                            catch (e_1_1) { e_1 = { error: e_1_1 }; }
                            finally {
                                try {
                                    if (_b && !_b.done && (_j = _a.return)) _j.call(_a);
                                }
                                finally { if (e_1) throw e_1.error; }
                            }
                        }
                        return [3, 7];
                    case 3:
                        if (!(type == "lease")) return [3, 5];
                        return [4, this.getQuesPickList("【賃貸管理】お問い合わせ", "'B0174'")];
                    case 4:
                        pickList = _o.sent();
                        if (pickList.records.length > 0) {
                            try {
                                for (_e = tslib_1.__values(pickList.records), _f = _e.next(); !_f.done; _f = _e.next()) {
                                    quesObj = _f.value;
                                    optionList = quesObj.optionProjectQuestionnaireQuestions__r;
                                    if (optionList && optionList.records.length > 0) {
                                        try {
                                            for (_g = (e_4 = void 0, tslib_1.__values(optionList.records)), _h = _g.next(); !_h.done; _h = _g.next()) {
                                                opObj = _h.value;
                                                option = { value: opObj.OptionItemNumber__c, label: opObj.OptionItemValue__c };
                                                if (quesObj.QuestionNumber__c == "B0174") {
                                                    yearTypeList.push(option);
                                                }
                                            }
                                        }
                                        catch (e_4_1) { e_4 = { error: e_4_1 }; }
                                        finally {
                                            try {
                                                if (_h && !_h.done && (_m = _g.return)) _m.call(_g);
                                            }
                                            finally { if (e_4) throw e_4.error; }
                                        }
                                    }
                                }
                            }
                            catch (e_3_1) { e_3 = { error: e_3_1 }; }
                            finally {
                                try {
                                    if (_f && !_f.done && (_l = _e.return)) _l.call(_e);
                                }
                                finally { if (e_3) throw e_3.error; }
                            }
                        }
                        return [3, 7];
                    case 5:
                        if (!(type == "present")) return [3, 7];
                        projectNumber = req.body.params.projectNumber;
                        presentQuery = "SELECT Id,QuestionNumber__c, QuestionContent__c,QuestionnaireAccount__r.Project__r.PropertyMail__c,\n        (select Id,OptionItemValue__c,OptionItemNumber__c from optionProjectQuestionnaireQuestions__r\n          where OptionItemDeleteFlag__c=false order by OptionItemOrderNumber__c)\n        FROM ProjectQuestionnaireQuestions__c where QuestionnaireAccount__r.ProjectNumber__c='" + projectNumber + "'";
                        return [4, this.sfdcService.query(presentQuery)];
                    case 6:
                        presentRes = _o.sent();
                        if (presentRes.totalSize > 0) {
                            presentInfo = presentRes;
                        }
                        _o.label = 7;
                    case 7:
                        resultJson = {
                            memberInfo: selectRresult,
                            newProjectList: newProjectList,
                            yearTypeList: yearTypeList,
                            presentInfo: presentInfo
                        };
                        res.json(resultJson);
                        return [3, 9];
                    case 8:
                        error_1 = _o.sent();
                        res.json("NG");
                        return [3, 9];
                    case 9: return [2];
                }
            });
        }); };
        this.sendIntroduce = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var state, bocType, projectType, selectQues, quesResult, contentDetail, projectNumber, fromMail, createBc, detail, dateTime, jsonData;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        state = req.body.params.state;
                        bocType = req.body.params.bocType;
                        projectType = "提携企業用紹介状発行";
                        if (bocType == "request") {
                            projectType = "提携企業用紹介状発行(資料請求)紹介状発行(提携企業)";
                        }
                        selectQues = "select ProjectNumber__c ,BcMembershipCreation__c,Project__r.PropertyMail__c,\n    (select QuestionNumber__c,QuestionContent__c from Questionnaire__r) from ProjectQuestionnaire__c\n     where OppRecordType__c ='\u53CD\u97FF'\n     and ProjectType__c = '" + projectType + "' and Project__r.ProjectCode__c ='" + state.projectCode + "'";
                        return [4, this.sfdcService.query(selectQues)];
                    case 1:
                        quesResult = _a.sent();
                        contentDetail = {};
                        projectNumber = "";
                        fromMail = "";
                        createBc = false;
                        if (quesResult.totalSize > 0) {
                            projectNumber = quesResult.records[0].ProjectNumber__c;
                            createBc = quesResult.records[0].BcMembershipCreation__c;
                            fromMail = quesResult.records[0].Project__r.PropertyMail__c;
                            detail = "{";
                            detail += '"A0001":"' + state.seiKanji + '",';
                            detail += '"A0002":"' + state.meiKanji + '",';
                            detail += '"A0003":"' + state.seiKana + '",';
                            detail += '"A0004":"' + state.meiKana + '",';
                            detail += '"A0013":"' + state.txtTel.value + '",';
                            detail += '"A0014":"' + state.txtMail + '",';
                            detail += '"C0062":"' + state.txtCompany + '",';
                            detail += '"B0220":"' + state.txtCooperationCounter + '",';
                            if (bocType == "request") {
                                detail += "\"B0005\":{\"" + this.getOptionKey(state.gender.value, state.gender.pickList) + "\":\"\"},";
                                detail +=
                                    '"A0016":"' + state.birthdayYear.value + "-" + state.birthdayMonth.value + "-" + state.birthdayDay.value + '",';
                                detail += "\"A0005\":{\"" + this.getOptionKey(state.abroadType.value, state.abroadType.pickList) + "\":\"\"},";
                                detail += '"A0006":"' + state.zipCode1.value + state.zipCode2.value + '",';
                                detail += '"A0007":"' + state.prefecture + '",';
                                detail += '"A0008":"' + state.city + '",';
                                detail += '"A0009":"' + state.town + '",';
                                detail += '"A0010":"' + state.chome + '",';
                                detail += '"A0011":"' + state.building + '",';
                            }
                            detail = detail.substring(0, detail.length - 1);
                            detail += "}";
                            contentDetail = JSON.parse(detail);
                        }
                        dateTime = this.getSystemDate();
                        jsonData = {
                            ProjectNumber: projectNumber,
                            AnswerCode: "",
                            CreateBC: createBc,
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
        this.sendInquiry = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var state, memberMyPageId, bocType, projectType, contactBody, quesResult, contentDetail, projectNumber, fromMail, createBc, dateTime, detail, textDeail, textDeail, jsonData;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        state = req.body.params.state;
                        memberMyPageId = state.memberMyPageId;
                        bocType = req.body.params.bocType;
                        projectType = "特典：ご入居までのご契約物件についてのお問い合わせ";
                        contactBody = "●ご入居までのご契約物件についてのお問い合わせ";
                        if (bocType == "etc") {
                            projectType = "特典：brilliaオーナーズクラブに関するお問い合わせ";
                            contactBody = "●brilliaオーナーズクラブに関するお問い合わせ";
                        }
                        else if (bocType == "reply") {
                            projectType = "特典：BOC問い合わせ";
                            contactBody = "●BOC問い合わせ";
                        }
                        else if (bocType == "lease") {
                            projectType = "【賃貸管理】お問い合わせ";
                        }
                        else if (bocType == "buy") {
                            projectType = "特典：東京建物分譲物件ご紹介特典(新）";
                            contactBody = "●東京建物分譲物件ご紹介特典";
                        }
                        return [4, this.getProjectQuesForBoc(projectType)];
                    case 1:
                        quesResult = _a.sent();
                        contentDetail = {};
                        projectNumber = "";
                        fromMail = "";
                        createBc = false;
                        dateTime = this.getSystemDate();
                        if (quesResult.totalSize > 0) {
                            projectNumber = quesResult.records[0].ProjectNumber__c;
                            createBc = quesResult.records[0].BcMembershipCreation__c;
                            fromMail = quesResult.records[0].Project__r.PropertyMail__c;
                            detail = "{";
                            if (bocType == "reply") {
                                textDeail = state.inquiryDetail.value.replaceAll("\n", "\\n");
                                textDeail = textDeail.replaceAll("\r\n", "\\n");
                                detail += '"B0193":"' + textDeail + '",';
                                detail += '"A0040":"' + state.contactName.value + '",';
                                detail += '"A0041":"' + state.contactNameKana.value + '",';
                                detail += '"B0182":"' + state.projectName.value + " " + state.contractRoomNum.value + '",';
                                detail += '"A0013":"' + state.phoneNumber.value + '",';
                                detail += '"A0014":"' + state.contractMail.value + '",';
                                contactBody += "\\n●問合せ内容：\\n" + textDeail +
                                    "\\n●お名前：" + state.contactName.value +
                                    "\\n●フリガナ：" + state.contactNameKana.value +
                                    "\\n●ご契約物件・部屋番号：" + state.projectName.value + " " + state.contractRoomNum.value +
                                    "\\n●メールアドレス：" + state.contractMail.value +
                                    "\\n●ご連絡先電話番号：" + state.phoneNumber.value;
                                detail += '"B0178":"' + contactBody + '",';
                            }
                            else if (bocType == "etc" || bocType == "bukken") {
                                textDeail = state.inquiryDetail.replaceAll("\n", "\\n");
                                textDeail = textDeail.replaceAll("\r\n", "\\n");
                                detail += '"A0040":"' + state.seiKanji + "  " + state.meiKanji + '",';
                                detail += '"A0041":"' + state.seiKana + "  " + state.meiKana + '",';
                                detail += '"B0182":"' + state.contractRoomNum.label + '",';
                                detail += '"A0013":"' + state.phoneNumber.value + '",';
                                detail += '"A0014":"' + state.contractMail.value + '",';
                                detail += '"B0193":"' + textDeail + '",';
                                contactBody += "\\n●問合せ内容：\\n" + textDeail +
                                    "\\n●お名前：" + state.seiKanji + "  " + state.meiKanji +
                                    "\\n●フリガナ：" + state.seiKana + "  " + state.meiKana +
                                    "\\n●ご契約物件・部屋番号：" + state.phoneNumber.value +
                                    "\\n●メールアドレス：" + state.contractMail.value +
                                    "\\n●ご連絡先電話番号：" + state.phoneNumber.value;
                                detail += '"B0178":"' + contactBody + '",';
                            }
                            else if (bocType == "lease") {
                                detail += '"A0001":"' + state.seiKanji + '",';
                                detail += '"A0002":"' + state.meiKanji + '",';
                                detail += '"A0003":"' + state.seiKana + '",';
                                detail += '"A0004":"' + state.meiKana + '",';
                                detail += '"B0182":"' + state.contractRoomNum.label + '",';
                                detail += '"A0013":"' + state.phoneNumber.value + '",';
                                detail += '"A0014":"' + state.contractMail.value + '",';
                                detail += '"A0006":"' + state.zipCode1 + state.zipCode2 + '",';
                                detail += '"A0007":"' + state.prefecture + '",';
                                detail += '"A0008":"' + state.city + '",';
                                detail += '"A0009":"' + state.town + '",';
                                detail += '"A0010":"' + state.chome + '",';
                                detail += '"A0011":"' + state.building + '",';
                                detail += '"B0169":"' + state.shipAddress + '",';
                                detail += '"B0183":"' + state.targetProject + '",';
                                detail += '"B0184":"' + state.projectType + '",';
                                detail += '"B0185":"' + state.projectTypeDetail + '",';
                                detail += '"B0186":"' + state.projectTypeName + '",';
                                detail += '"B0187":"' + state.roomCnt + '",';
                                detail += '"B0188":"' + state.floorCnt + '",';
                                detail += '"B0189":"' + state.leaseProject + ' ' + state.leaseProjectAddress + '",';
                                detail += '"B0190":"' + state.buildYearName.value + state.buildYear + '年' + state.buildMonth + '月",';
                                detail += '"B0191":"' + state.management + '",';
                                detail += '"B0192":"' + state.otherRemark + '",';
                            }
                            else if (bocType == "buy") {
                                detail += '"A0001":"' + state.seiKanji + '",';
                                detail += '"A0002":"' + state.meiKanji + '",';
                                detail += '"A0003":"' + state.seiKana + '",';
                                detail += '"A0004":"' + state.meiKana + '",';
                                detail += '"B0182":"' + state.contractRoomNum.label + '",';
                                detail += '"A0013":"' + state.phoneNumber.value + '",';
                                detail += '"A0014":"' + state.contractMail.value + '",';
                                detail += '"A0006":"' + state.shippingAddress.zipCode + '",';
                                detail += '"A0007":"' + state.shippingAddress.prefectures + '",';
                                detail += '"A0008":"' + state.shippingAddress.city + '",';
                                detail += '"A0009":"' + state.shippingAddress.towns + '",';
                                detail += '"A0010":"' + state.shippingAddress.address + '",';
                                detail += '"A0011":"' + state.shippingAddress.room + '",';
                                detail += '"A0040":"' + state.seiKanjiBuy.value + "  " + state.meiKanjiBuy.value + '",';
                                detail += '"A0041":"' + state.seiKanaBuy.value + "  " + state.meiKanaBuy.value + '",';
                                detail += '"B0194":{"' + this.getOptionKey(state.projectBuy.value, state.projectBuy.pickList) + '":""},';
                                detail += '"A0043":"' + state.zipCode1.value + state.zipCode2.value + '",';
                                detail += '"A0044":"' + state.prefecture.value + '",';
                                detail += '"A0045":"' + state.city.value + '",';
                                detail += '"A0046":"' + state.town.value + '",';
                                detail += '"A0047":"' + state.chome.value + '",';
                                detail += '"A0048":"' + state.building.value + '",';
                                detail += '"A0049":"' + state.phoneNumberBuy.value + '",';
                                contactBody += "\\n●お客様情報：" +
                                    "\\n●お名前：" + state.seiKanji + state.meiKanji +
                                    "\\n●フリガナ：" + state.seiKana + state.meiKana +
                                    "\\n●ご契約物件・部屋番号：" + state.contractRoomNum.label +
                                    "\\n●メールアドレス：" + state.contractMail.value +
                                    "\\n●ご連絡先電話番号：" + state.phoneNumber.value +
                                    "\\n●住所：" + state.shippingAddress.zipCode + state.shippingAddress.prefectures + state.shippingAddress.city +
                                    state.shippingAddress.towns + state.shippingAddress.address + state.shippingAddress.room +
                                    "\\n\\n●ご購入検討のお客様" +
                                    "\\n●お名前：" + state.seiKanjiBuy.value + "  " + state.meiKanjiBuy.value +
                                    "\\n●フリガナ：" + state.seiKanaBuy.value + "  " + state.meiKanaBuy.value +
                                    "\\n●ご検討物件名：" + state.projectBuy.value +
                                    "\\n●住所：" + state.zipCode1.value + state.zipCode2.value + state.prefecture.value + state.city.value +
                                    state.town.value + state.chome.value + state.building.value +
                                    "\\n●ご連絡先電話番号：" + state.phoneNumberBuy.value;
                                detail += '"B0178":"' + contactBody + '",';
                            }
                            detail = detail.substring(0, detail.length - 1);
                            detail += "}";
                            contentDetail = JSON.parse(detail);
                        }
                        jsonData = {
                            ProjectNumber: projectNumber,
                            AnswerCode: "",
                            CreateBC: createBc,
                            AnswerDate: dateTime,
                            Create: dateTime,
                            BOCID: memberMyPageId,
                            Content: contentDetail
                        };
                        console.log(JSON.stringify(jsonData));
                        this.sendMail(fromMail, state.toMail, "フォーム投稿", jsonData);
                        res.send("ok");
                        return [2];
                }
            });
        }); };
        this.sendPresent = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var state, memberMyPageId, contentDetail, projectNumber, fromMail, dateTime, detail, contactBody, timeArry, dateFmt, title, jsonData;
            return tslib_1.__generator(this, function (_a) {
                state = req.body.params.state;
                memberMyPageId = state.memberMyPageId;
                contentDetail = {};
                projectNumber = state.projectNumber;
                fromMail = state.projectMail;
                dateTime = this.getSystemDate();
                detail = "{";
                contactBody = "\\n●お客様情報のご入力：" +
                    "\\n　お名前：" + state.seiKanji + state.meiKanji +
                    "\\n　フリガナ：" + state.seiKana + state.meiKana +
                    "\\n　ご契約物件・部屋番号：" + state.contractRoomNum.label +
                    "\\n　メールアドレス：" + state.contractMail.value +
                    "\\n　ご連絡が可能な電話番号：" + state.phoneNumber.value +
                    "\\n　希望の賞品：" + state.present.label +
                    "\\n　郵便番号：" + state.shippingAddress.zipCode +
                    "\\n　都道府県：" + state.shippingAddress.prefectures +
                    "\\n　市区名：" + state.shippingAddress.city +
                    "\\n　町村名：" + state.shippingAddress.towns +
                    "\\n　丁目番地：" + state.shippingAddress.address +
                    "\\n　建物名・部屋番号：" + state.shippingAddress.room;
                detail += '"B0178":"' + contactBody + '",';
                timeArry = dateTime.split(" ")[0].split("-");
                dateFmt = timeArry[0] + timeArry[1] + timeArry[2];
                title = "BOC\u5FDC\u52DF\u8005\u3010" + state.present.label + "\u3011\u30D7\u30EC\u30BC\u30F3\u30C8\u5FDC\u52DF\u30D5\u30A9\u30FC\u30E0" + dateFmt;
                detail += '"B0221":"' + title + '",';
                detail = detail.substring(0, detail.length - 1);
                detail += "}";
                contentDetail = JSON.parse(detail);
                jsonData = {
                    ProjectNumber: projectNumber,
                    AnswerCode: "",
                    CreateBC: false,
                    AnswerDate: dateTime,
                    Create: dateTime,
                    BOCID: memberMyPageId,
                    Content: contentDetail
                };
                console.log(JSON.stringify(jsonData));
                this.sendMail(fromMail, state.toMail, "フォーム投稿", jsonData);
                res.send("ok");
                return [2];
            });
        }); };
        this.passwordChange = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var state, projectType, quesResult, contentDetail, projectNumber, fromMail, createBc, detail, dateTime, jsonData;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        state = req.body.params.state;
                        projectType = "ID、パスワード忘れ 会員登録状況の確認";
                        return [4, this.getProjectQuesForBoc(projectType)];
                    case 1:
                        quesResult = _a.sent();
                        contentDetail = {};
                        projectNumber = "";
                        fromMail = "";
                        createBc = false;
                        if (quesResult.totalSize > 0) {
                            projectNumber = quesResult.records[0].ProjectNumber__c;
                            createBc = quesResult.records[0].BcMembershipCreation__c;
                            fromMail = quesResult.records[0].Project__r.PropertyMail__c;
                            detail = "{";
                            detail += '"A0001":"' + state.seiKanji.value + '",';
                            detail += '"A0002":"' + state.meiKanji.value + '",';
                            detail += '"A0014":"' + state.memberMail.value + '"';
                            detail += "}";
                            contentDetail = JSON.parse(detail);
                        }
                        dateTime = this.getSystemDate();
                        jsonData = {
                            ProjectNumber: projectNumber,
                            AnswerCode: "",
                            CreateBC: createBc,
                            AnswerDate: dateTime,
                            Create: dateTime,
                            BOCID: state.memberPageId,
                            Content: contentDetail
                        };
                        console.log(JSON.stringify(jsonData));
                        this.sendMail(fromMail, state.toMail, "フォーム投稿", jsonData);
                        res.send("ok");
                        return [2];
                }
            });
        }); };
        this.checkOpp = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var query, quesResult;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        query = req.body.params.query;
                        return [4, this.sfdcService.query(query)];
                    case 1:
                        quesResult = _a.sent();
                        res.send(quesResult);
                        return [2];
                }
            });
        }); };
        this.checkProfile = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var memberMyPageId, oppId, selectQuery, selectRresult, checkRes;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        memberMyPageId = req.body.params.memberMyPageId;
                        oppId = req.body.params.oppId;
                        selectQuery = "SELECT Id FROM Opportunity WHERE MemberMyPageID__c = '" + memberMyPageId + "' and Id <> '" + oppId + "'";
                        return [4, this.sfdcService.query(selectQuery)];
                    case 1:
                        selectRresult = _a.sent();
                        checkRes = "OK";
                        if (selectRresult.totalSize > 0) {
                            checkRes = "NG";
                        }
                        res.send(checkRes);
                        return [2];
                }
            });
        }); };
        this.getProfileInfo = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var memberMyPageId, selectQuery, selectRresult, hopeDeliveryList, shippingRelationshipList, relationshipList, ownerTypeList, viewModelList, genderList, selectUser, quesResult, _a, _b, quesObj, optionList, _c, _d, opObj, option, resultJson, err_1;
            var e_5, _e, e_6, _f;
            return tslib_1.__generator(this, function (_g) {
                switch (_g.label) {
                    case 0:
                        _g.trys.push([0, 7, , 8]);
                        memberMyPageId = req.body.params.memberMyPageId;
                        selectQuery = "SELECT Id,AccountId,Account.ProjectCode__c,ContactName__c,ContactNameKana__c,Contact__r.LastName,Contact__r.FirstName, " +
                            "Contact__r.Birthdate,PersonName__c,AddressRelationship__c,RequestMailMemberNewsletter__c," +
                            "ZipCode__c,Prefectures__c,Towns__c,RoomNumPropertyName__c,City__c,Address__c,MemberMyPagePwd__c,MemberMyPageID__c" +
                            ",(select Id,MailAddress__c,ContactPhoneNumber__c,MainFlag__c,FamilyName__c,FamilyNameKana__c," +
                            " FirstName__c,FirstNameKana__c,Birthday__c,Gender__c,ViewModel__c " +
                            " from contactEmailOpportunity__r)" +
                            ",(select Id,BocProjectName__c,BocRoom__c,OwnershipForm__c,SubYearNum__c,SubYear__c,SubMonth__c,RegisterDate__c from BOC__r )" +
                            " FROM Opportunity WHERE MemberMyPageID__c = '" + memberMyPageId + "' and MemberWithdrawalDate__c = null and RecordType.DeveloperName='BOC' and RecordType.SobjectType='Opportunity'";
                        return [4, this.sfdcService.query(selectQuery)];
                    case 1:
                        selectRresult = _g.sent();
                        hopeDeliveryList = [];
                        shippingRelationshipList = [];
                        return [4, this.sfdcService.getPickList("ContactEmailList__c", "ContactRelation__c")];
                    case 2:
                        relationshipList = _g.sent();
                        return [4, this.sfdcService.getPickList("BOCMemberRoom__c", "OwnershipForm__c")];
                    case 3:
                        ownerTypeList = _g.sent();
                        return [4, this.sfdcService.getPickList("ContactEmailList__c", "ViewModel__c")];
                    case 4:
                        viewModelList = _g.sent();
                        return [4, this.sfdcService.getPickList("ContactEmailList__c", "Gender__c")];
                    case 5:
                        genderList = _g.sent();
                        selectUser = "select Id ,QuestionNumber__c,\n        (select Id, OptionItemValue__c,OptionItemNumber__c from optionProjectQuestionnaireQuestions__r\n          where OptionItemDeleteFlag__c=false\n          order by MainItemOrderNumber__c,OptionItemOrderNumber__c )\n        from ProjectQuestionnaireQuestions__c\n        where QuestionnaireAccount__r.ProjectType__c ='BOC\u60C5\u5831\u5909\u66F4/\u767B\u9332\u5185\u5BB9\u5909\u66F4' and QuestionDeleteFlag__c = false\n        and QuestionnaireAccount__r.Project__r.ProjectType__c='BOC\u7528'\n        and QuestionNumber__c in ('B0029', 'B0168') order by OrderNumber__c ";
                        return [4, this.sfdcService.query(selectUser)];
                    case 6:
                        quesResult = _g.sent();
                        if (quesResult.records.length > 0) {
                            try {
                                for (_a = tslib_1.__values(quesResult.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                    quesObj = _b.value;
                                    optionList = quesObj.optionProjectQuestionnaireQuestions__r;
                                    if (optionList && optionList.records.length > 0) {
                                        try {
                                            for (_c = (e_6 = void 0, tslib_1.__values(optionList.records)), _d = _c.next(); !_d.done; _d = _c.next()) {
                                                opObj = _d.value;
                                                option = { value: opObj.OptionItemNumber__c, label: opObj.OptionItemValue__c };
                                                if (quesObj.QuestionNumber__c == "B0168") {
                                                    hopeDeliveryList.push(option);
                                                }
                                                if (quesObj.QuestionNumber__c == "B0029") {
                                                    shippingRelationshipList.push(option);
                                                }
                                            }
                                        }
                                        catch (e_6_1) { e_6 = { error: e_6_1 }; }
                                        finally {
                                            try {
                                                if (_d && !_d.done && (_f = _c.return)) _f.call(_c);
                                            }
                                            finally { if (e_6) throw e_6.error; }
                                        }
                                    }
                                }
                            }
                            catch (e_5_1) { e_5 = { error: e_5_1 }; }
                            finally {
                                try {
                                    if (_b && !_b.done && (_e = _a.return)) _e.call(_a);
                                }
                                finally { if (e_5) throw e_5.error; }
                            }
                        }
                        resultJson = {
                            memberInfo: selectRresult,
                            hopeDeliveryList: hopeDeliveryList,
                            shippingRelationshipList: shippingRelationshipList,
                            relationshipList: relationshipList,
                            ownerTypeList: ownerTypeList,
                            viewModelList: viewModelList,
                            genderList: genderList
                        };
                        res.json(resultJson);
                        return [3, 8];
                    case 7:
                        err_1 = _g.sent();
                        res.json("NG");
                        return [3, 8];
                    case 8: return [2];
                }
            });
        }); };
        this.sendUpdateProfile = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var state, memberMyPageId, updateType, subjectDetail, quesResult, contentDetail, bcCreateFlg, propertyContents, familyInformation, projectNumber, fromMail, detail, editDetail, newVal, _a, _b, opObj, hopeDelivery, _c, _d, opObj, _e, _f, room, roomInfo, _g, _h, familyObj, familyInfo, familyObj, familyInfo, dateTime, jsonData;
            var e_7, _j, e_8, _k, e_9, _l, e_10, _m;
            return tslib_1.__generator(this, function (_o) {
                switch (_o.label) {
                    case 0:
                        state = req.body.params.state;
                        memberMyPageId = state.memberMyPageId;
                        updateType = req.body.params.updateType;
                        subjectDetail = "BOC情報変更/登録内容変更";
                        return [4, this.getProjectQuesForBoc(subjectDetail)];
                    case 1:
                        quesResult = _o.sent();
                        contentDetail = {};
                        bcCreateFlg = false;
                        propertyContents = [];
                        familyInformation = [];
                        projectNumber = "";
                        fromMail = "";
                        if (quesResult.totalSize > 0) {
                            projectNumber = quesResult.records[0].ProjectNumber__c;
                            bcCreateFlg = quesResult.records[0].BcMembershipCreation__c;
                            fromMail = quesResult.records[0].Project__r.PropertyMail__c;
                            detail = "{";
                            editDetail = "";
                            if (updateType == "update") {
                                detail += '"B0179":"' + state.memberMyPageId + '",';
                                editDetail = this.getEditDeatail(editDetail, "会員マイページID", state.memberMyPageId, state.memberInfoOld.MemberMyPageID__c);
                                memberMyPageId = state.memberInfoOld.MemberMyPageID__c;
                                detail += '"B0180":"' + state.memberMyPagePwd + '",';
                                detail += '"A0014":"' + state.memberMail.value + '",';
                                editDetail = this.getEditDeatail(editDetail, "メールアドレス", state.memberMail.value, state.memberInfoOld.memberMail);
                                detail += '"A0006":"' + state.zipCode1 + state.zipCode2 + '",';
                                editDetail = this.getEditDeatail(editDetail, "送付先住所 郵便番号", state.zipCode1 + state.zipCode2, state.memberInfoOld.ZipCode__c);
                                detail += '"A0007":"' + state.prefecturesList.value + '",';
                                editDetail = this.getEditDeatail(editDetail, "送付先住所 都道府県", state.prefecturesList.value, state.memberInfoOld.Prefectures__c);
                                detail += '"A0008":"' + state.cityList.value + '",';
                                editDetail = this.getEditDeatail(editDetail, "送付先住所 市区名", state.cityList.value, state.memberInfoOld.City__c);
                                detail += '"A0009":"' + state.shippingTowns + '",';
                                editDetail = this.getEditDeatail(editDetail, "送付先住所 町村名", state.shippingTowns, state.memberInfoOld.Towns__c);
                                detail += '"A0010":"' + state.shippingAddress + '",';
                                editDetail = this.getEditDeatail(editDetail, "送付先住所 丁目番地", state.shippingAddress, state.memberInfoOld.Address__c);
                                detail += '"A0011":"' + state.shippingRoomNumPropertyName + '",';
                                editDetail = this.getEditDeatail(editDetail, "送付先住所 建物名・部屋番号", state.shippingRoomNumPropertyName, state.memberInfoOld.RoomNumPropertyName__c);
                                detail += '"B0169":"' + state.shippingPersonName + '",';
                                editDetail = this.getEditDeatail(editDetail, "送付先住所 送り先宛名", state.shippingPersonName, state.memberInfoOld.PersonName__c);
                                detail += '"B0029":{"' + state.shippingRelation + '":""},';
                                newVal = state.shippingRelation;
                                try {
                                    for (_a = tslib_1.__values(state.shippingRelationshipList), _b = _a.next(); !_b.done; _b = _a.next()) {
                                        opObj = _b.value;
                                        if (opObj.value == state.shippingRelation) {
                                            newVal = opObj.label;
                                        }
                                    }
                                }
                                catch (e_7_1) { e_7 = { error: e_7_1 }; }
                                finally {
                                    try {
                                        if (_b && !_b.done && (_j = _a.return)) _j.call(_a);
                                    }
                                    finally { if (e_7) throw e_7.error; }
                                }
                                editDetail = this.getEditDeatail(editDetail, "送付先住所 送付先続柄", newVal, state.memberInfoOld.AddressRelationship__c);
                                detail += '"A0013":"' + state.phoneNumber + '",';
                                editDetail = this.getEditDeatail(editDetail, "電話番号", state.phoneNumber, state.memberInfoOld.phoneNumber);
                                detail += '"B0168":{"' + state.hopeDelivery + '":""},';
                                hopeDelivery = state.hopeDelivery;
                                try {
                                    for (_c = tslib_1.__values(state.hopeDeliveryList), _d = _c.next(); !_d.done; _d = _c.next()) {
                                        opObj = _d.value;
                                        if (opObj.value == state.hopeDelivery) {
                                            hopeDelivery = opObj.label;
                                        }
                                    }
                                }
                                catch (e_8_1) { e_8 = { error: e_8_1 }; }
                                finally {
                                    try {
                                        if (_d && !_d.done && (_k = _c.return)) _k.call(_c);
                                    }
                                    finally { if (e_8) throw e_8.error; }
                                }
                                editDetail = this.getEditDeatail(editDetail, "送付物の希望", hopeDelivery, state.memberInfoOld.RequestMailMemberNewsletter__c);
                                if (state.bocRoomList.length > 0) {
                                    try {
                                        for (_e = tslib_1.__values(state.bocRoomList), _f = _e.next(); !_f.done; _f = _e.next()) {
                                            room = _f.value;
                                            roomInfo = {
                                                PropertyID: "",
                                                OwnershipForm: ""
                                            };
                                            roomInfo.PropertyID = room.Id;
                                            if (room.OwnershipForm__c != null) {
                                                roomInfo.OwnershipForm = room.OwnershipForm__c;
                                            }
                                            if (room.OwnershipForm__c != room.OwnershipFormOld) {
                                                editDetail = this.getEditDeatail(editDetail, "所有形態(" + room.BocProjectName__c + ")", room.OwnershipForm__c, room.OwnershipFormOld);
                                            }
                                            propertyContents.push(roomInfo);
                                        }
                                    }
                                    catch (e_9_1) { e_9 = { error: e_9_1 }; }
                                    finally {
                                        try {
                                            if (_f && !_f.done && (_l = _e.return)) _l.call(_e);
                                        }
                                        finally { if (e_9) throw e_9.error; }
                                    }
                                }
                                if (state.familyList.length > 0) {
                                    try {
                                        for (_g = tslib_1.__values(state.familyList), _h = _g.next(); !_h.done; _h = _g.next()) {
                                            familyObj = _h.value;
                                            if (familyObj.deleteFlg) {
                                                familyInfo = {
                                                    FamilyID: familyObj.Id,
                                                    DeleteFlag: true,
                                                    FamilyName: "",
                                                    FirstName: "",
                                                    FamilyNameKana: "",
                                                    FirstNameKana: "",
                                                    Birthdate: null,
                                                    Gender: "",
                                                    Mail: "",
                                                    BrowsingModel: "",
                                                    Relationship: ""
                                                };
                                                editDetail += "家族削除：" + familyObj.FamilyName__c + familyObj.FirstName__c + ";";
                                                familyInformation.push(familyInfo);
                                            }
                                        }
                                    }
                                    catch (e_10_1) { e_10 = { error: e_10_1 }; }
                                    finally {
                                        try {
                                            if (_h && !_h.done && (_m = _g.return)) _m.call(_g);
                                        }
                                        finally { if (e_10) throw e_10.error; }
                                    }
                                }
                                detail += '"B0181":"' + editDetail + '",';
                                detail = detail.substring(0, detail.length - 1);
                                detail += "}";
                                contentDetail = JSON.parse(detail);
                            }
                            else {
                                familyObj = state.newFamily;
                                editDetail = "家族追加：" + familyObj.txtFamilySeiKanji + familyObj.txtFamilyMeiKanji + ";";
                                detail += '"B0181":"' + editDetail + '"}';
                                contentDetail = JSON.parse(detail);
                                familyInfo = {
                                    FamilyID: "",
                                    DeleteFlag: false,
                                    FamilyName: familyObj.txtFamilySeiKanji,
                                    FirstName: familyObj.txtFamilyMeiKanji,
                                    FamilyNameKana: familyObj.txtFamilySeikana,
                                    FirstNameKana: familyObj.txtFamilyMeikana,
                                    Birthdate: familyObj.txtFamilyBirthdayYear +
                                        "-" +
                                        familyObj.txtFamilyBirthdayMonth +
                                        "-" +
                                        familyObj.txtFamilyBirthdayDay,
                                    Gender: familyObj.txtFamilyGender,
                                    Mail: familyObj.txtFamilyMailAddress,
                                    BrowsingModel: familyObj.txtFamilyViewType,
                                    Relationship: familyObj.txtFamilyRelation
                                };
                                familyInformation.push(familyInfo);
                            }
                        }
                        dateTime = this.getSystemDate();
                        jsonData = {
                            ProjectNumber: projectNumber,
                            AnswerCode: "",
                            CreateBC: bcCreateFlg,
                            AnswerDate: dateTime,
                            Create: dateTime,
                            BOCID: memberMyPageId,
                            Content: contentDetail,
                            PropertyContents: propertyContents,
                            FamilyInformation: familyInformation
                        };
                        console.log(JSON.stringify(jsonData));
                        this.sendMail(fromMail, state.toMail, "フォーム投稿", jsonData);
                        res.send("ok");
                        return [2];
                }
            });
        }); };
        this.initJoin = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var hopeDeliveryList, addressArry, projetIdArray, addressQuery, projectAddressList, _a, _b, projectAdd, prefecture, city, projectId, selectUser, quesResult, ownerTypeList, contactRelationList, viewModelList, ownerClassificationList, shippingRelationshipList, genderList, joinClusList, mailFormatList, ownerRouteList, _c, _d, quesObj, optionList, _e, _f, opObj, option, resultJson, err_2;
            var e_11, _g, e_12, _h, e_13, _j;
            return tslib_1.__generator(this, function (_k) {
                switch (_k.label) {
                    case 0:
                        _k.trys.push([0, 5, , 6]);
                        hopeDeliveryList = [];
                        addressArry = [];
                        projetIdArray = [{}];
                        addressQuery = "select PrefectureResAddress__c PrefectureResAddress, MunicipalResAddress__c MunicipalResAddress\n        from Building__c where PrefectureResAddress__c <> '' and MunicipalResAddress__c <> ''\n        group by PrefectureResAddress__c, MunicipalResAddress__c";
                        return [4, this.sfdcService.query(addressQuery)];
                    case 1:
                        projectAddressList = _k.sent();
                        if (projectAddressList.totalSize > 0) {
                            try {
                                for (_a = tslib_1.__values(projectAddressList.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                    projectAdd = _b.value;
                                    prefecture = projectAdd.PrefectureResAddress;
                                    city = projectAdd.MunicipalResAddress;
                                    projectId = "";
                                    addressArry.push({ prefecture: prefecture, city: city, projectId: projectId });
                                }
                            }
                            catch (e_11_1) { e_11 = { error: e_11_1 }; }
                            finally {
                                try {
                                    if (_b && !_b.done && (_g = _a.return)) _g.call(_a);
                                }
                                finally { if (e_11) throw e_11.error; }
                            }
                        }
                        selectUser = "select Id ,QuestionNumber__c,\n        (select Id, OptionItemValue__c,OptionItemNumber__c from optionProjectQuestionnaireQuestions__r\n          where OptionItemDeleteFlag__c=false\n          order by MainItemOrderNumber__c,OptionItemOrderNumber__c )\n        from ProjectQuestionnaireQuestions__c\n        where QuestionnaireAccount__r.ProjectType__c ='BOC\u65B0\u898F\u5165\u4F1A\u30D5\u30A9\u30FC\u30E0' and QuestionDeleteFlag__c = false\n        and QuestionnaireAccount__r.Project__r.ProjectType__c='BOC\u7528'\n        and QuestionNumber__c in ('B0005','B0170', 'B0029', 'B0168','B0165', 'B0166', 'B0072', 'B0167') order by OrderNumber__c ";
                        return [4, this.sfdcService.query(selectUser)];
                    case 2:
                        quesResult = _k.sent();
                        return [4, this.sfdcService.getPickList("BOCMemberRoom__c", "OwnershipForm__c")];
                    case 3:
                        ownerTypeList = _k.sent();
                        return [4, this.sfdcService.getPickList("ContactEmailList__c", "ContactRelation__c")];
                    case 4:
                        contactRelationList = _k.sent();
                        viewModelList = [];
                        ownerClassificationList = [];
                        shippingRelationshipList = [];
                        genderList = [];
                        joinClusList = [];
                        mailFormatList = [];
                        ownerRouteList = [];
                        if (quesResult.records.length > 0) {
                            try {
                                for (_c = tslib_1.__values(quesResult.records), _d = _c.next(); !_d.done; _d = _c.next()) {
                                    quesObj = _d.value;
                                    optionList = quesObj.optionProjectQuestionnaireQuestions__r;
                                    if (optionList && optionList.records.length > 0) {
                                        try {
                                            for (_e = (e_13 = void 0, tslib_1.__values(optionList.records)), _f = _e.next(); !_f.done; _f = _e.next()) {
                                                opObj = _f.value;
                                                option = { value: opObj.OptionItemNumber__c, label: opObj.OptionItemValue__c };
                                                if (quesObj.QuestionNumber__c == "B0005") {
                                                    genderList.push(option);
                                                }
                                                if (quesObj.QuestionNumber__c == "B0165") {
                                                    ownerClassificationList.push(option);
                                                }
                                                if (quesObj.QuestionNumber__c == "B0168") {
                                                    hopeDeliveryList.push(option);
                                                }
                                                if (quesObj.QuestionNumber__c == "B0170") {
                                                    joinClusList.push(option);
                                                }
                                                if (quesObj.QuestionNumber__c == "B0029") {
                                                    shippingRelationshipList.push(option);
                                                }
                                                if (quesObj.QuestionNumber__c == "B0072") {
                                                    ownerRouteList.push(option);
                                                }
                                                if (quesObj.QuestionNumber__c == "B0166") {
                                                    mailFormatList.push(option);
                                                }
                                                if (quesObj.QuestionNumber__c == "B0167") {
                                                    viewModelList.push(option);
                                                }
                                            }
                                        }
                                        catch (e_13_1) { e_13 = { error: e_13_1 }; }
                                        finally {
                                            try {
                                                if (_f && !_f.done && (_j = _e.return)) _j.call(_e);
                                            }
                                            finally { if (e_13) throw e_13.error; }
                                        }
                                    }
                                }
                            }
                            catch (e_12_1) { e_12 = { error: e_12_1 }; }
                            finally {
                                try {
                                    if (_d && !_d.done && (_h = _c.return)) _h.call(_c);
                                }
                                finally { if (e_12) throw e_12.error; }
                            }
                        }
                        resultJson = {
                            hopeDeliveryList: hopeDeliveryList,
                            shippingRelationshipList: shippingRelationshipList,
                            ownerTypeList: ownerTypeList,
                            viewModelList: viewModelList,
                            ownerClassificationList: ownerClassificationList,
                            genderList: genderList,
                            contactRelationList: contactRelationList,
                            addressArry: addressArry,
                            projetIdArray: projetIdArray,
                            recognitionList: joinClusList,
                            mailFormatList: mailFormatList,
                            ownerRouteList: ownerRouteList
                        };
                        res.json(resultJson);
                        return [3, 6];
                    case 5:
                        err_2 = _k.sent();
                        res.json("NG");
                        return [3, 6];
                    case 6: return [2];
                }
            });
        }); };
        this.sendBocJoin = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var state, subjectDetail, quesResult, contentDetail, bcCreateFlg, propertyContents, familyInformation, projectNumber, fromMail, roomId, roomQuery, roomResult, detail, viewType, _a, _b, obj, roomInfo, _c, _d, familyObj, familyInfo, dateTime, jsonData;
            var e_14, _e, e_15, _f;
            return tslib_1.__generator(this, function (_g) {
                switch (_g.label) {
                    case 0:
                        state = req.body.params.state;
                        subjectDetail = "BOC新規入会フォーム";
                        return [4, this.getProjectQuesForBoc(subjectDetail)];
                    case 1:
                        quesResult = _g.sent();
                        contentDetail = {};
                        bcCreateFlg = false;
                        propertyContents = [];
                        familyInformation = [];
                        projectNumber = "";
                        fromMail = "";
                        roomId = "";
                        if (!(state.projectId && state.propertyRoom)) return [3, 3];
                        roomQuery = "select Id from Room__c where Project__c='" + state.projectId + "' and Name='" + state.propertyRoom + "'";
                        return [4, this.sfdcService.query(roomQuery)];
                    case 2:
                        roomResult = _g.sent();
                        if (roomResult.totalSize > 0) {
                            roomId = roomResult.records[0].Id;
                        }
                        _g.label = 3;
                    case 3:
                        if (quesResult.totalSize > 0) {
                            projectNumber = quesResult.records[0].ProjectNumber__c;
                            bcCreateFlg = quesResult.records[0].BcMembershipCreation__c;
                            fromMail = quesResult.records[0].Project__r.PropertyMail__c;
                            detail = "{";
                            detail += '"A0001":"' + state.memberSeiKanji + '",';
                            detail += '"A0002":"' + state.memberMeiKanji + '",';
                            detail += '"A0003":"' + state.memberSeiKana + '",';
                            detail += '"A0004":"' + state.memberMeiKana + '",';
                            detail += '"B0165":{"' + state.ownerClassification.value + '":""},';
                            detail +=
                                '"A0016":"' + state.memberBirthdayYear + "-" + state.memberBirthdayMonth + "-" + state.memberBirthdayDay + '",';
                            detail += '"B0005":{"' + state.memberGender.value + '":""},';
                            detail += '"A0013":"' + state.phoneNumber.value + '",';
                            detail += '"A0014":"' + state.memberMail.value + '",';
                            detail += '"B0166":{"' + state.mailFormat.value + '":""},';
                            viewType = state.viewType;
                            try {
                                for (_a = tslib_1.__values(state.viewModelList), _b = _a.next(); !_b.done; _b = _a.next()) {
                                    obj = _b.value;
                                    if (state.viewType == obj.label) {
                                        viewType = obj.value;
                                    }
                                }
                            }
                            catch (e_14_1) { e_14 = { error: e_14_1 }; }
                            finally {
                                try {
                                    if (_b && !_b.done && (_e = _a.return)) _e.call(_a);
                                }
                                finally { if (e_14) throw e_14.error; }
                            }
                            detail += '"B0167":{"' + viewType + '":""},';
                            detail += '"B0168":{"' + state.hopeDelivery.value + '":""},';
                            detail += '"A0006":"' + state.zipCode1 + state.zipCode2 + '",';
                            detail += '"A0007":"' + state.shippingPrefecture + '",';
                            detail += '"A0008":"' + state.shippingCity + '",';
                            detail += '"A0009":"' + state.shippingTowns + '",';
                            detail += '"A0010":"' + state.shippingAddress + '",';
                            detail += '"A0011":"' + state.shippingRoomNumPropertyName + '",';
                            detail += '"B0169":"' + state.shippingPersonName + '",';
                            detail += '"B0029":{"' + state.shippingRelation.value + '":""},';
                            detail += '"B0170":{"' + state.recognitionMedium.value + '":"' + state.recognitionMedium.otherVal + '"},';
                            if (state.recognitionMedium.otherVal) {
                                detail += '"B0216":"' + state.recognitionMedium.otherVal + '",';
                            }
                            roomInfo = {
                                PropertyID: "",
                                Property: state.projectId,
                                OwnershipForm: state.ownerType,
                                Room: roomId,
                                PropertyInput: state.propertyNameInput,
                                RoomInput: state.propertyRoomInput,
                                OwnershipRoute: state.ownershipRoute.label
                            };
                            propertyContents.push(roomInfo);
                            detail = detail.substring(0, detail.length - 1);
                            detail += "}";
                            contentDetail = JSON.parse(detail);
                            if (state.newFamilyList.length > 0) {
                                try {
                                    for (_c = tslib_1.__values(state.newFamilyList), _d = _c.next(); !_d.done; _d = _c.next()) {
                                        familyObj = _d.value;
                                        familyInfo = {
                                            FamilyID: "",
                                            DeleteFlag: false,
                                            FamilyName: familyObj.txtFamilySeiKanji,
                                            FirstName: familyObj.txtFamilyMeiKanji,
                                            FamilyNameKana: familyObj.txtFamilySeikana,
                                            FirstNameKana: familyObj.txtFamilyMeikana,
                                            Birthdate: familyObj.txtFamilyBirthdayYear +
                                                "-" +
                                                familyObj.txtFamilyBirthdayMonth +
                                                "-" +
                                                familyObj.txtFamilyBirthdayDay,
                                            Gender: familyObj.txtFamilyGender,
                                            Mail: familyObj.txtFamilyMailAddress,
                                            BrowsingModel: familyObj.txtFamilyViewType,
                                            Relationship: familyObj.txtFamilyRelation
                                        };
                                        familyInformation.push(familyInfo);
                                    }
                                }
                                catch (e_15_1) { e_15 = { error: e_15_1 }; }
                                finally {
                                    try {
                                        if (_d && !_d.done && (_f = _c.return)) _f.call(_c);
                                    }
                                    finally { if (e_15) throw e_15.error; }
                                }
                            }
                        }
                        dateTime = this.getSystemDate();
                        jsonData = {
                            ProjectNumber: projectNumber,
                            AnswerCode: "",
                            CreateBC: bcCreateFlg,
                            CampaignCode: state.campainCode,
                            AnswerDate: dateTime,
                            Create: dateTime,
                            BOCID: "",
                            Content: contentDetail,
                            PropertyContents: propertyContents,
                            FamilyInformation: familyInformation
                        };
                        console.log(JSON.stringify(jsonData));
                        this.sendMail(fromMail, state.toMail, "フォーム投稿", jsonData);
                        res.send("ok");
                        return [2];
                }
            });
        }); };
        this.initWithdraw = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var withdrawReson, pickList, _a, _b, quesObj, optionList, _c, _d, opObj, option, err_3;
            var e_16, _e, e_17, _f;
            return tslib_1.__generator(this, function (_g) {
                switch (_g.label) {
                    case 0:
                        _g.trys.push([0, 2, , 3]);
                        withdrawReson = [];
                        return [4, this.getQuesPickList("BOC退会フォーム", "'B0096'")];
                    case 1:
                        pickList = _g.sent();
                        if (pickList.records.length > 0) {
                            try {
                                for (_a = tslib_1.__values(pickList.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                    quesObj = _b.value;
                                    optionList = quesObj.optionProjectQuestionnaireQuestions__r;
                                    if (optionList && optionList.records.length > 0) {
                                        try {
                                            for (_c = (e_17 = void 0, tslib_1.__values(optionList.records)), _d = _c.next(); !_d.done; _d = _c.next()) {
                                                opObj = _d.value;
                                                option = { value: opObj.OptionItemNumber__c, label: opObj.OptionItemValue__c };
                                                if (quesObj.QuestionNumber__c == "B0096") {
                                                    withdrawReson.push(option);
                                                }
                                            }
                                        }
                                        catch (e_17_1) { e_17 = { error: e_17_1 }; }
                                        finally {
                                            try {
                                                if (_d && !_d.done && (_f = _c.return)) _f.call(_c);
                                            }
                                            finally { if (e_17) throw e_17.error; }
                                        }
                                    }
                                }
                            }
                            catch (e_16_1) { e_16 = { error: e_16_1 }; }
                            finally {
                                try {
                                    if (_b && !_b.done && (_e = _a.return)) _e.call(_a);
                                }
                                finally { if (e_16) throw e_16.error; }
                            }
                        }
                        res.json(withdrawReson);
                        return [3, 3];
                    case 2:
                        err_3 = _g.sent();
                        res.json("NG");
                        return [3, 3];
                    case 3: return [2];
                }
            });
        }); };
        this.sendWithdraw = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var state, quesResult, contentDetail, bcCreateFlg, projectNumber, fromMail, detail, withdrawDate, dateTime, jsonData;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        state = req.body.params.state;
                        return [4, this.getProjectQuesForBoc("BOC退会フォーム")];
                    case 1:
                        quesResult = _a.sent();
                        contentDetail = {};
                        bcCreateFlg = false;
                        projectNumber = "";
                        fromMail = "";
                        if (quesResult.totalSize > 0) {
                            projectNumber = quesResult.records[0].ProjectNumber__c;
                            bcCreateFlg = quesResult.records[0].BcMembershipCreation__c;
                            fromMail = quesResult.records[0].Project__r.PropertyMail__c;
                            detail = "{";
                            detail += '"B0096":{"' + state.withdrawReson.value + '":""},';
                            withdrawDate = state.withdrawDate.replaceAll("/", "-");
                            detail += '"B0175":"' + withdrawDate + '"';
                            detail += "}";
                            console.log(detail);
                            contentDetail = JSON.parse(detail);
                        }
                        dateTime = this.getSystemDate();
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
                        return [2];
                }
            });
        }); };
        this.getProjectByAddress = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var propertyPrefecture, propertyCity, projectQuery, projectList, projetIdArray, _a, _b, project;
            var e_18, _c;
            return tslib_1.__generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        propertyPrefecture = req.body.params.propertyPrefecture;
                        propertyCity = req.body.params.propertyCity;
                        projectQuery = "select Project__c Project, Project__r.Name Name\n      from Building__c where PrefectureResAddress__c = '" + propertyPrefecture + "' and MunicipalResAddress__c = '" + propertyCity + "'\n      group by Project__c, Project__r.Name";
                        return [4, this.sfdcService.query(projectQuery)];
                    case 1:
                        projectList = _d.sent();
                        projetIdArray = [];
                        if (projectList.totalSize > 0) {
                            try {
                                for (_a = tslib_1.__values(projectList.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                    project = _b.value;
                                    projetIdArray.push({ proId: project.Project, proName: project.Name });
                                }
                            }
                            catch (e_18_1) { e_18 = { error: e_18_1 }; }
                            finally {
                                try {
                                    if (_b && !_b.done && (_c = _a.return)) _c.call(_a);
                                }
                                finally { if (e_18) throw e_18.error; }
                            }
                        }
                        res.json(projetIdArray);
                        return [2];
                }
            });
        }); };
        this.getAddressMaster = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var quesResult, type, prefectures, selectUser, zipcode, addressResult, addressInfo, pref, addDeatai, selectCity, rest, rest;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        quesResult = null;
                        type = req.body.params.type;
                        prefectures = req.body.params.prefectures;
                        selectUser = "select Prefectures__c Prefectures from AddressMst__c group by Prefectures__c";
                        if (!(type == "searchCity")) return [3, 2];
                        selectUser =
                            "select City__c City from AddressMst__c where Prefectures__c = '" + prefectures + "' group by City__c";
                        return [4, this.sfdcService.query(selectUser)];
                    case 1:
                        quesResult = _a.sent();
                        res.json(quesResult);
                        return [3, 6];
                    case 2:
                        if (!(type == "searchByZipcode")) return [3, 4];
                        zipcode = req.body.params.zipCode;
                        selectUser =
                            "select Prefectures__c,City__c,Address__c,Towns__c,PostCode__c from AddressMst__c where PostCode__c = '" +
                                zipcode +
                                "' ";
                        return [4, this.sfdcService.query(selectUser)];
                    case 3:
                        addressResult = _a.sent();
                        if (addressResult.totalSize > 0) {
                            addressInfo = addressResult.records[0];
                            pref = addressInfo.Prefectures__c;
                            addDeatai = "";
                            if (addressResult.totalSize == 1) {
                                addDeatai = addressInfo.Address__c;
                            }
                            selectCity = [];
                            "select City__c City from AddressMst__c where Prefectures__c = '" + pref + "' group by City__c";
                            rest = {
                                prefectures: addressInfo.Prefectures__c,
                                city: addressInfo.City__c,
                                towns: addressInfo.Towns__c,
                                addDeatai: addDeatai,
                                cityList: [],
                                resList: addressResult.records
                            };
                            res.json(rest);
                        }
                        else {
                            rest = {
                                resList: addressResult.records
                            };
                            res.json(rest);
                        }
                        return [3, 6];
                    case 4: return [4, this.sfdcService.query(selectUser)];
                    case 5:
                        quesResult = _a.sent();
                        res.json(quesResult);
                        _a.label = 6;
                    case 6: return [2];
                }
            });
        }); };
        this.getProjectInfo = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var proCode, selectQuery, quesResult, selectUser, pickList, genderList, overseaList, _a, _b, quesObj, optionList, _c, _d, opObj, option, resultJson;
            var e_19, _e, e_20, _f;
            return tslib_1.__generator(this, function (_g) {
                switch (_g.label) {
                    case 0:
                        proCode = req.body.params.proCode;
                        selectQuery = "select Id, Name from Account where ProjectCode__c='" + proCode + "'";
                        return [4, this.sfdcService.query(selectQuery)];
                    case 1:
                        quesResult = _g.sent();
                        selectUser = "select Id ,QuestionNumber__c,\n      (select Id, OptionItemValue__c,OptionItemNumber__c from optionProjectQuestionnaireQuestions__r\n        where OptionItemDeleteFlag__c=false\n        order by OptionItemOrderNumber__c )\n      from ProjectQuestionnaireQuestions__c\n      where QuestionnaireAccount__r.Project__r.ProjectCode__c='" + proCode + "'\n      and QuestionnaireAccount__r.ProjectType__c='\u63D0\u643A\u4F01\u696D\u7528\u7D39\u4ECB\u72B6\u767A\u884C(\u8CC7\u6599\u8ACB\u6C42)\u7D39\u4ECB\u72B6\u767A\u884C(\u63D0\u643A\u4F01\u696D)'\n      and QuestionNumber__c in ('B0005','A0005')";
                        return [4, this.sfdcService.query(selectUser)];
                    case 2:
                        pickList = _g.sent();
                        genderList = [];
                        overseaList = [];
                        if (pickList.records.length > 0) {
                            try {
                                for (_a = tslib_1.__values(pickList.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                    quesObj = _b.value;
                                    optionList = quesObj.optionProjectQuestionnaireQuestions__r;
                                    if (optionList && optionList.records.length > 0) {
                                        try {
                                            for (_c = (e_20 = void 0, tslib_1.__values(optionList.records)), _d = _c.next(); !_d.done; _d = _c.next()) {
                                                opObj = _d.value;
                                                option = { value: opObj.OptionItemNumber__c, label: opObj.OptionItemValue__c };
                                                if (quesObj.QuestionNumber__c == "B0005") {
                                                    genderList.push(option);
                                                }
                                                if (quesObj.QuestionNumber__c == "A0005") {
                                                    overseaList.push(option);
                                                }
                                            }
                                        }
                                        catch (e_20_1) { e_20 = { error: e_20_1 }; }
                                        finally {
                                            try {
                                                if (_d && !_d.done && (_f = _c.return)) _f.call(_c);
                                            }
                                            finally { if (e_20) throw e_20.error; }
                                        }
                                    }
                                }
                            }
                            catch (e_19_1) { e_19 = { error: e_19_1 }; }
                            finally {
                                try {
                                    if (_b && !_b.done && (_e = _a.return)) _e.call(_a);
                                }
                                finally { if (e_19) throw e_19.error; }
                            }
                        }
                        resultJson = {
                            projectInfo: quesResult,
                            genderList: genderList,
                            overseaList: overseaList
                        };
                        res.json(resultJson);
                        return [2];
                }
            });
        }); };
        this.login = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var username, password, selectUser, quesResult, loginRes, _a, _b, userInfo, loginUserInfo, _c, _d, bocPro;
            var e_21, _e, e_22, _f;
            return tslib_1.__generator(this, function (_g) {
                switch (_g.label) {
                    case 0:
                        username = req.body.params.state.userName;
                        password = req.body.params.state.userPassword;
                        selectUser = "select Id, MemberMyPageID__c ,MemberMyPagePwd__c,ContactName__c,\n        (select Project__r.ProjectCode__c from BOC__r)\n      from Opportunity where MemberMyPageID__c = '" + username + "'\n      and MemberWithdrawalDate__c = null and RecordType.DeveloperName='BOC' and RecordType.SobjectType='Opportunity'";
                        return [4, this.sfdcService.query(selectUser)];
                    case 1:
                        quesResult = _g.sent();
                        loginRes = {
                            result: "NG",
                            recordID: "",
                            memberMyPageID: username,
                            memberMyPagePwd: password,
                            ckNameUserInfoLogin: ""
                        };
                        if (quesResult.totalSize > 0) {
                            try {
                                for (_a = tslib_1.__values(quesResult.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                    userInfo = _b.value;
                                    console.log(userInfo.MemberMyPagePwd__c + "==" + password);
                                    if (userInfo.MemberMyPagePwd__c == password) {
                                        loginRes.result = "OK";
                                        loginRes.recordID = userInfo.Id;
                                        loginUserInfo = userInfo.ContactName__c;
                                        if (userInfo.BOC__r.totalSize > 0) {
                                            try {
                                                for (_c = (e_22 = void 0, tslib_1.__values(userInfo.BOC__r.records)), _d = _c.next(); !_d.done; _d = _c.next()) {
                                                    bocPro = _d.value;
                                                    loginUserInfo += "(|)" + bocPro.Project__r.ProjectCode__c;
                                                }
                                            }
                                            catch (e_22_1) { e_22 = { error: e_22_1 }; }
                                            finally {
                                                try {
                                                    if (_d && !_d.done && (_f = _c.return)) _f.call(_c);
                                                }
                                                finally { if (e_22) throw e_22.error; }
                                            }
                                        }
                                        loginRes.ckNameUserInfoLogin = loginUserInfo;
                                        break;
                                    }
                                }
                            }
                            catch (e_21_1) { e_21 = { error: e_21_1 }; }
                            finally {
                                try {
                                    if (_b && !_b.done && (_e = _a.return)) _e.call(_a);
                                }
                                finally { if (e_21) throw e_21.error; }
                            }
                        }
                        res.send(JSON.stringify(loginRes));
                        return [2];
                }
            });
        }); };
        this.getProjectQuesForBoc = function (subjectDetail) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var selectQues, quesResult;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        selectQues = "select ProjectNumber__c ,BcMembershipCreation__c,Project__r.PropertyMail__c,\n    (select QuestionNumber__c,QuestionContent__c from Questionnaire__r) from ProjectQuestionnaire__c\n     where OppRecordType__c ='BOC' and Project__r.ProjectType__c='BOC\u7528'\n     and ProjectType__c = '" + subjectDetail + "'";
                        return [4, this.sfdcService.query(selectQues)];
                    case 1:
                        quesResult = _a.sent();
                        return [2, quesResult];
                }
            });
        }); };
        this.getQuesPickList = function (projectType, qesNumList) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var selectUser, quesResult;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        selectUser = "select Id ,QuestionNumber__c,\n      (select Id, OptionItemValue__c,OptionItemNumber__c from optionProjectQuestionnaireQuestions__r\n        where OptionItemDeleteFlag__c=false\n        order by MainItemOrderNumber__c,OptionItemOrderNumber__c )\n      from ProjectQuestionnaireQuestions__c\n      where QuestionnaireAccount__r.ProjectType__c ='" + projectType + "' and QuestionDeleteFlag__c = false\n      and QuestionnaireAccount__r.Project__r.ProjectType__c='BOC\u7528'\n      and QuestionNumber__c in (" + qesNumList + ") order by OrderNumber__c ";
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
    BocCommonController.prototype.intializeRoutes = function () {
        this.router.post("/boclogin", this.login);
        this.router.post("/initInquiry", this.initInquiry);
        this.router.post("/sendInquiry", this.sendInquiry);
        this.router.post("/getProfileInfo", this.getProfileInfo);
        this.router.post("/sendUpdateProfile", this.sendUpdateProfile);
        this.router.post("/initJoin", this.initJoin);
        this.router.post("/sendBocJoin", this.sendBocJoin);
        this.router.post("/getAddressMaster", this.getAddressMaster);
        this.router.post("/getProjectByAddress", this.getProjectByAddress);
        this.router.post("/checkProfile", this.checkProfile);
        this.router.post("/passwordChange", this.passwordChange);
        this.router.post("/checkOpp", this.checkOpp);
        this.router.post("/initWithdraw", this.initWithdraw);
        this.router.post("/sendWithdraw", this.sendWithdraw);
        this.router.post("/getProjectInfo", this.getProjectInfo);
        this.router.post("/sendIntroduce", this.sendIntroduce);
        this.router.post("/sendPresent", this.sendPresent);
    };
    BocCommonController.prototype.getEditDeatail = function (editDetail, fieldName, newVal, oldVal) {
        if (newVal != oldVal) {
            editDetail += fieldName + "更新：" + this.nullToString(oldVal) + "-->" + this.nullToString(newVal) + ";";
        }
        return editDetail;
    };
    BocCommonController.prototype.sendMail = function (fromMail, toMail, subjectDetail, jsonData) {
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
    BocCommonController.prototype.getSystemDate = function () {
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
    BocCommonController.prototype.nullToString = function (target) {
        var resStr = "";
        if (target != undefined && target != null && target != "") {
            resStr = target;
        }
        return resStr;
    };
    BocCommonController.prototype.getOptionKey = function (label, targetList) {
        var e_23, _a;
        var resVal = label;
        try {
            for (var targetList_1 = tslib_1.__values(targetList), targetList_1_1 = targetList_1.next(); !targetList_1_1.done; targetList_1_1 = targetList_1.next()) {
                var opObj = targetList_1_1.value;
                if (label == opObj.label) {
                    resVal = opObj.value;
                }
            }
        }
        catch (e_23_1) { e_23 = { error: e_23_1 }; }
        finally {
            try {
                if (targetList_1_1 && !targetList_1_1.done && (_a = targetList_1.return)) _a.call(targetList_1);
            }
            finally { if (e_23) throw e_23.error; }
        }
        return resVal;
    };
    return BocCommonController;
}());
exports.BocCommonController = BocCommonController;
//# sourceMappingURL=boc.common.controller.js.map