/* eslint-disable @typescript-eslint/camelcase */
import { Router, Request, Response, NextFunction } from "express"
import { BaseController } from "./base.controller"
import { SfdcService } from "../service/sfdc.service"
import { HttpResult } from "../model/httpResult"
import { OutputConfig } from "../model/outputConfig"
import { QuestionnaireService } from "../service/questionnaire.service"

export class FilesController implements BaseController {
  public router: Router

  constructor() {
    this.router = Router()
    this.intializeRoutes()
  }

  public intializeRoutes() {
    this.router.post("/api/questionnaire", this.getQuestionnaire)
    this.router.get("/api/test/:Id", this.testDownLoadFile)
  }

  testDownLoadFile = async (req: Request, res: Response) => {
    const optConfig: OutputConfig = {
      accountId: "test",
      opportunityId: "test",
      userId: "test",
      opportunityName: ""
    }

    optConfig.opportunityId = req.params.Id

    const ars = new QuestionnaireService(optConfig)
    await ars.init()
    console.log("出力成功")
    // const fileBuffer = await ars.getFileBuffer()
    // res.set("Content-disposition", "attachment; filename=" + "demo.xlsx")
    // res.set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    // return res.end(fileBuffer)
    return res.send("ok")
  }

  /**
   * 調査票出力
   */
  getQuestionnaire = async (req: Request, res: Response, next: NextFunction) => {
    let httpRes: HttpResult = { status: 0, data: null, msg: "調査票ファイル作成中" }
    try {
      const optConfig: OutputConfig = req.body

      this.gennaretaExcel("調査票", optConfig, next)

      // const ars = new AnalysReportService(analysisRptConfig, next)
      // await ars.init()
      // const fileBuffer = await ars.getFileBuffer()
      // res.set("Content-disposition", "attachment; filename=" + "analysis.xlsx")
      // res.set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      // return res.end(fileBuffer)
      return res.json(httpRes)
    } catch (error) {
      console.error(error)
      next(error)
      return res.json(error)
    }
  }

  /**
   *
   * @param oppId 調査票スペースID
   * @param fileName ファイル名
   * @param analysisRptConfig 調査票条件
   * @param next
   */
  gennaretaExcel = async (fileName: string, analysisRptConfig: OutputConfig, next: NextFunction) => {
    const ars = new QuestionnaireService(analysisRptConfig)
    let error
    await ars.init().catch(err => {
      error = err
      console.error(err)
      next(err)
    })
    // メール通知
    const sfdcService = new SfdcService()
    // エラーが存在する場合
    if (error) {
      sfdcService.sendChatter(
        `
        ${fileName}が作成失敗しました。システム管理者にご連絡下さい。
        エラーメッセージ:
        ${error}
        プロジェクト:
      `,
        [analysisRptConfig.userId],
        `${sfdcService.getDomainUrl()}/lightning/r/Account/${analysisRptConfig.accountId}/view`
      )
      return
    }

    sfdcService.sendChatter(
      `
      ${fileName}が作成できました。プロジェクト関連するFleekDrive内にファイルをご確認ください。
      ファイル名: ${analysisRptConfig.accountId}
      プロジェクト:
    `,
      [analysisRptConfig.userId],
      `${sfdcService.getDomainUrl()}/lightning/r/Account/${analysisRptConfig.accountId}/view`
    )
  }
}
