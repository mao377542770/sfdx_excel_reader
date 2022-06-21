"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.BocLoginController = void 0;
var tslib_1 = require("tslib");
var express_1 = require("express");
var sfdc_service_1 = require("../service/sfdc.service");
var BocLoginController = (function () {
    function BocLoginController() {
        var _this = this;
        this.sfdcService = new sfdc_service_1.SfdcService();
        this.SECRET = process.env.LOGIN_SECRET || "ITFORCE";
        this.login = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var username, password, selectUser, quesResult, _a, _b, userInfo;
            var e_1, _c;
            return tslib_1.__generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        username = req.body.params.state.userName;
                        password = req.body.params.state.userPassword;
                        selectUser = "select Id, MemberMyPageID__c ,MemberMyPagePwd__c,ContactName__c "
                            + ",(select Project__r.ProjectCode__c from BOC__r)"
                            + " from Opportunity "
                            + " where MemberMyPageID__c ='" + username + "' and MemberWithdrawalDate__c = NULL";
                        return [4, this.sfdcService.query(selectUser)];
                    case 1:
                        quesResult = _d.sent();
                        if (quesResult.totalSize > 0) {
                            try {
                                for (_a = tslib_1.__values(quesResult.records), _b = _a.next(); !_b.done; _b = _a.next()) {
                                    userInfo = _b.value;
                                    console.log(userInfo.MemberMyPagePwd__c);
                                    console.log("console.log(password) " + userInfo.MemberMyPagePwd__c == password);
                                    if (userInfo.MemberMyPagePwd__c == password) {
                                        console.log("login ok");
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
                        }
                        res.send("OK");
                        return [2];
                }
            });
        }); };
        this.router = express_1.Router();
        this.intializeRoutes();
    }
    BocLoginController.prototype.intializeRoutes = function () {
    };
    return BocLoginController;
}());
exports.BocLoginController = BocLoginController;
//# sourceMappingURL=boclogin.controller.js.map