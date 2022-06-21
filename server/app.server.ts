import express, { NextFunction, Request, Response } from "express"
import { BaseController } from "./controller/base.controller"
import cors from "cors"
import Jwt from "jsonwebtoken"

// URLホワイトリスト
const whiteListUrl = ["/login"]

export default class App {
  public app: express.Application
  public port: number

  constructor(controllers: BaseController[], port: number) {
    this.app = express()
    this.port = port

    this.initializeMiddlewares()
    this.initializeControllers(controllers)
  }

  private initializeMiddlewares() {
    this.errorCatch()
    this.app.use(
      cors({
        origin: "*",
        credentials: true
      })
    )
    this.app.use((req, res, next) => {
      res.header("Access-Control-Allow-Origin", "*")
      res.header("Access-Control-Allow-Methods", "GET,PUT,POST,DELETE")
      res.header("Access-Control-Allow-Headers", "Content-Type")
      next()
    })
    this.app.use(express.urlencoded({ extended: false }))
    this.app.use(express.json())

    // this.app.use(history())
    // this.app.use(express.static("dist/client"))
  }

  private hasOneOf(str: string, arr: Array<string>) {
    return arr.some(item => item.includes(str))
  }

  private accessFilter() {
    // 中间件（类似拦截器，每次请求时候会走这里）
    this.app.all("*", (req, res, next) => {
      const path = req.path
      if (this.hasOneOf(path, whiteListUrl)) {
        next() // 認証必要ない
      } else {
        // リクエストのheaderからauthorizationを取得
        const token = req.headers.authorization
        if (!token) res.status(401).send("there is no token, please login")
        else {
          // 有token就判断使用JWT提供的方法，校验token是否正确
          // 第一个参数把获取到的token传入
          // 第二个参数是我们生成token时候传入的密钥，当然这个密钥可以抽离出一个文件每次从文件中获取密钥
          // 第三个参数是一个回调函数，第一个参数是错误信息，第二个参数是从token中解码出来的信息
          Jwt.verify(token, "abcd", (error, decode) => {
            if (error)
              res.send({
                // token错误
                code: 401,
                mes: "token error",
                data: {}
              })
            else {
              // token正确返回用户名，继续往下走
              console.log(decode)
              next()
            }
          })
        }
      }
    })
  }

  private initializeControllers(controllers: BaseController[]) {
    controllers.forEach(controller => {
      this.app.use("/", controller.router)
    })
  }

  private errorCatch() {
    this.app.use((err: any, req: Request, res: Response, next: NextFunction) => {
      console.error(err)
      res.status(500).send("System Error!")
    })
  }

  public listen() {
    this.app.listen(this.port, () => {
      console.log(`App listening on the port ${this.port}`)
    })
  }
}
