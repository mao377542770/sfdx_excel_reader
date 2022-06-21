"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var app_server_1 = tslib_1.__importDefault(require("./app.server"));
var files_controller_1 = require("./controller/files.controller");
var dotenv_1 = tslib_1.__importDefault(require("dotenv"));
var opportunity_controller_1 = require("./controller/opportunity.controller");
var boc_common_controller_1 = require("./controller/boc.common.controller");
var boc_questionary_controller_1 = require("./controller/boc.questionary.controller");
var boclogin_controller_1 = require("./controller/boclogin.controller");
var bc_common_controller_1 = require("./controller/bc.common.controller");
dotenv_1.default.config();
var port = process.env.PORT || 8080;
var app = new app_server_1.default([
    new files_controller_1.FilesController(),
    new opportunity_controller_1.OpportunityController(),
    new boc_common_controller_1.BocCommonController(),
    new boclogin_controller_1.BocLoginController(),
    new boc_questionary_controller_1.BocQuestionaryController(),
    new bc_common_controller_1.BcCommonController()
], port);
app.listen();
//# sourceMappingURL=app.js.map