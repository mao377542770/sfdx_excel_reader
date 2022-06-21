"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var express_1 = tslib_1.__importDefault(require("express"));
var cors_1 = tslib_1.__importDefault(require("cors"));
var connect_history_api_fallback_1 = tslib_1.__importDefault(require("connect-history-api-fallback"));
var jsonwebtoken_1 = tslib_1.__importDefault(require("jsonwebtoken"));
var whiteListUrl = ["/login"];
var App = (function () {
    function App(controllers, port) {
        this.app = express_1.default();
        this.port = port;
        this.initializeMiddlewares();
        this.initializeControllers(controllers);
    }
    App.prototype.initializeMiddlewares = function () {
        this.errorCatch();
        this.app.use(cors_1.default({
            origin: "*",
            credentials: true
        }));
        this.app.use(function (req, res, next) {
            res.header("Access-Control-Allow-Origin", "*");
            res.header("Access-Control-Allow-Methods", "GET,PUT,POST,DELETE");
            res.header("Access-Control-Allow-Headers", "Content-Type");
            next();
        });
        this.app.use(express_1.default.urlencoded({ extended: false }));
        this.app.use(express_1.default.json());
        this.app.use(connect_history_api_fallback_1.default());
        this.app.use(express_1.default.static("dist/client"));
    };
    App.prototype.hasOneOf = function (str, arr) {
        return arr.some(function (item) { return item.includes(str); });
    };
    App.prototype.accessFilter = function () {
        var _this = this;
        this.app.all("*", function (req, res, next) {
            var path = req.path;
            if (_this.hasOneOf(path, whiteListUrl)) {
                next();
            }
            else {
                var token = req.headers.authorization;
                if (!token)
                    res.status(401).send("there is no token, please login");
                else {
                    jsonwebtoken_1.default.verify(token, "abcd", function (error, decode) {
                        if (error)
                            res.send({
                                code: 401,
                                mes: "token error",
                                data: {}
                            });
                        else {
                            console.log(decode);
                            next();
                        }
                    });
                }
            }
        });
    };
    App.prototype.initializeControllers = function (controllers) {
        var _this = this;
        controllers.forEach(function (controller) {
            _this.app.use("/", controller.router);
        });
    };
    App.prototype.errorCatch = function () {
        this.app.use(function (err, req, res, next) {
            console.error(err);
            res.status(500).send("System Error!");
        });
    };
    App.prototype.listen = function () {
        var _this = this;
        this.app.listen(this.port, function () {
            console.log("App listening on the port " + _this.port);
        });
    };
    return App;
}());
exports.default = App;
//# sourceMappingURL=app.server.js.map