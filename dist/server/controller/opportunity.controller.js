"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.OpportunityController = void 0;
var tslib_1 = require("tslib");
var express_1 = require("express");
var sfdc_service_1 = require("../service/sfdc.service");
var OpportunityController = (function () {
    function OpportunityController() {
        var _this = this;
        this.sfdcService = new sfdc_service_1.SfdcService();
        this.getOpportunity = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var records;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4, this.sfdcService.getData()];
                    case 1:
                        records = _a.sent();
                        res.json(records);
                        return [2];
                }
            });
        }); };
        this.router = express_1.Router();
        this.intializeRoutes();
    }
    OpportunityController.prototype.intializeRoutes = function () {
        this.router.get("/api/opportunity", this.getOpportunity);
    };
    return OpportunityController;
}());
exports.OpportunityController = OpportunityController;
//# sourceMappingURL=opportunity.controller.js.map