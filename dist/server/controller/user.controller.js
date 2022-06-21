"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.UserController = void 0;
var tslib_1 = require("tslib");
var express_1 = require("express");
var sfdc_service_1 = require("../service/sfdc.service");
var jsonwebtoken_1 = tslib_1.__importDefault(require("jsonwebtoken"));
var UserController = (function () {
    function UserController() {
        var _this = this;
        this.sfdcService = new sfdc_service_1.SfdcService();
        this.SECRET = process.env.LOGIN_SECRET || "ITFORCE";
        this.login = function (req, res) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var _a, username, password;
            return tslib_1.__generator(this, function (_b) {
                _a = req.body, username = _a.username, password = _a.password;
                if (username === "admin" && password === "admin") {
                    res.send({
                        code: 200,
                        mes: "success",
                        data: {
                            token: jsonwebtoken_1.default.sign({ username: username }, this.SECRET, {
                                expiresIn: "1h"
                            })
                        }
                    });
                }
                return [2];
            });
        }); };
        this.router = express_1.Router();
        this.intializeRoutes();
    }
    UserController.prototype.intializeRoutes = function () {
        this.router.get("/login", this.login);
    };
    return UserController;
}());
exports.UserController = UserController;
//# sourceMappingURL=user.controller.js.map