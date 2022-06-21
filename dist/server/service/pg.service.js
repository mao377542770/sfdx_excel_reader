"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.PgService = void 0;
var tslib_1 = require("tslib");
var pg_1 = require("pg");
var dotenv_1 = tslib_1.__importDefault(require("dotenv"));
dotenv_1.default.config();
var DATABASE_URL = process.env.DATABASE_URL;
var conn = new pg_1.Client({
    connectionString: DATABASE_URL,
    ssl: {
        rejectUnauthorized: false
    }
});
conn.connect();
var PgService = (function () {
    function PgService() {
    }
    PgService.prototype.query = function (sql) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var ps;
            return tslib_1.__generator(this, function (_a) {
                console.log("[Heroku Sql]:" + sql);
                ps = new Promise(function (resolve, reject) {
                    PgService.conn.query(sql, function (err, res) {
                        if (err) {
                            reject(err);
                            throw err;
                        }
                        resolve(res);
                    });
                }).catch(function (err) {
                    throw err;
                });
                return [2, ps];
            });
        });
    };
    PgService.conn = conn;
    return PgService;
}());
exports.PgService = PgService;
//# sourceMappingURL=pg.service.js.map