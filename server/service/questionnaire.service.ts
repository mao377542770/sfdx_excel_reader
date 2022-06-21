/* eslint-disable @typescript-eslint/camelcase */
import ExcelJS, { Cell, Worksheet } from "exceljs"
import { OutputConfig } from "../model/outputConfig"
import { SfdcService } from "./sfdc.service"
import { Opportunity } from "../model/opportunity"
import { Account } from "../model/account"
import { AccountQuestionnaire__c } from "../model/accountQuestionnaire__c"

interface CellInfo {
  sfdcId: string
  rowNum: number
}

//調査票改行ソース
export class QuestionnaireService {
  private workbook: ExcelJS.Workbook
  private sheetLsit: ExcelJS.Worksheet[]
  public worksheet?: ExcelJS.Worksheet
  public sfdcService: SfdcService
  private outputConfig: OutputConfig
  private PageInfo: any

  // 商談調査票
  public static opptyTyouFileList = [
    "POSITION_CONTRACTOR__C",
    "NAME_CONTRACTOR__C",
    "POSTCODE_CONTRACTOR_RES__C",
    "COMPANYADRESS2_CONTRACTOR_RES__C",
    "COMPANYADRESS3_CONTRACTOR_RES__C",
    "COMPANYADRESS5_CONTRACTOR_RES__C",
    "COMPANYADRESS6_CONTRACTOR_RES__C",
    "COMPANYADRESS7_CONTRACTOR_RES__C",
    "COMPANYADRESS8_CONTRACTOR_RES__C",
    "COMPANYNAME_CONTRATOR_RES__C",
    "POSTION_CONTRACTOR_RES__C",
    "NAME_CONTRACTOR_RES__C",
    "E_MAIL_CONTRACTOR_RES__C",
    "TEL_CONTRACTOR_RES__C",
    "FAX_CONTRACTOR_RES__C"
  ]

  // 重要場所項目のリスト
  public static LocationFiledList = [
    "ID",
    "FACILITYNAME__C",
    "FACILITYNAMEKANA__C",
    "SUPPLYPOINTIDENTIFICATIONNUMBER__C",
    "KEISOKUDATE__C",
    "SUPPLYSTARTDATE__C",
    "SUPPLYENDDATE__C",
    "POSTCODE_CONTRACTOR_RES__C",
    "COMPANYADRESS2_CONTRACTOR_RES__C",
    "COMPANYADRESS3_CONTRACTOR_RES__C",
    "COMPANYADRESS5_CONTRACTOR_RES__C",
    "COMPANYADRESS6_CONTRACTOR_RES__C",
    "COMPANYADRESS7_CONTRACTOR_RES__C",
    "COMPANYADRESS8_CONTRACTOR_RES__C",
    "DEMANDERWINDOWDEPARTMENT__C",
    "DEMANDERWINDOWNAME__C",
    "DEMANDERWINDOWPHONE__C",
    "INVOICERESPONSIBLERDEPARTMENT__C",
    "INVOICERESPONSIBLERNAME__C",
    "INVOICERESPONSIBLERPHONE__C",
    "CONTRACTNUMBEROFCURRENTRETAILER__C",
    "INVOICEWINDOWDEPARTMENT__C",
    "INVOICEWINDOWNAME__C",
    "INVOICEWINDOWPHONE__C",
    "INVOICEWINDOWEMAIL__C",
    "INVOICEDESCRIBEDINFOPOSTCODE__C",
    "INVOICEDESCRIBEDINFOADDRESS2__C",
    "INVOICEDESCRIBEDINFOADDRESS3__C",
    "INVOICEDESCRIBEDINFOADDRESS4__C",
    "INVOICEDESCRIBEDINFOADDRESS5__C",
    "INVOICEDESCRIBEDINFODESTINATION__C",
    "INVOICEDESCRIBEDINFONAME__C",
    "INVOICEPROCESS__C",
    "BANKNAME__C",
    "BANKTYPE__C",
    "BANKBRANCHNAME__C",
    "BANKACCOUNTNUMBER__C",
    "BANKACCOUNTNAME__C"
  ]

  constructor(outputConfig: OutputConfig) {
    this.workbook = new ExcelJS.Workbook()
    this.sheetLsit = []
    this.sfdcService = new SfdcService()
    this.outputConfig = outputConfig
  }

  public async init() {
    const fileStream = await this.sfdcService.downFileByDocumentId("0690w000001XfkAAAS")

    //週報マスタのテンプレートファイルを取得する
    this.workbook = await this.workbook.xlsx.read(fileStream)

    //合計情報リセット
    this.sheetLsit = this.workbook.worksheets
    //XX分析表
    await this.setAnasislyTable().catch(err => {
      throw err
    })
  }

  public async getFileBuffer() {
    return await this.workbook.xlsx.writeBuffer()
  }

  //XX分析表 の設定
  private async setAnasislyTable() {
    // 調査票(新)
    const questionnaireSheet = this.getWorkSheetByName(`【ご提出下さい】調査票 (新)`)
    // 請求単位
    const requestSheet = this.getWorkSheetByName(`【ご提出下さい】請求単位 (新)`)

    let promiseJobList: Promise<any>[] = []

    if (questionnaireSheet) {
      promiseJobList.push(this.setTitleSheet(questionnaireSheet))
      promiseJobList.push(this.setQuestionnaireSheet(questionnaireSheet))
    }

    await Promise.all(promiseJobList).catch(err => {
      throw err
    })
  }

  private async setTitleSheet(sheet: ExcelJS.Worksheet) {
    // 需要家
    const sql = `SELECT Id,CompanyName_Contractor__c,CompanyNameKana_Contractor__c,Address_Contract__c FROM Account WHERE ID = '0010w00000xh6O9AAI'`
    const accRes = await this.sfdcService.query<Account>(sql)
    let acc: Account
    if (!accRes || accRes.totalSize === 0) acc = accRes.records[0]

    // 商談調査票
    const aqSql = `SELECT ${SfdcService.getQueryFileds(
      QuestionnaireService.opptyTyouFileList
    )} FROM AccountQuestionnaire__c WHERE Id = 'a280w000000L2zuAAC'`
    const aqRes = await this.sfdcService.query<AccountQuestionnaire__c>(aqSql)
    let aq: AccountQuestionnaire__c
    if (!aqRes || aqRes.totalSize === 0) aq = aqRes.records[0]

    sheet.getRows(4, 5)?.forEach(row => {
      row.eachCell((cell, colNumber) => {
        if (cell && cell.value) {
          const valStr = cell.value.toString()
          const regStr = valStr.match("{.*}")
          if (regStr) {
            const filedApiName = regStr[0].substring(1, regStr[0].length - 1)
            if (acc && acc[filedApiName]) {
              // 需要家値を設定する
              cell.value = acc[filedApiName]
            } else if (aq && aq[filedApiName]) {
              // 商談調査票
              cell.value = aq[filedApiName]
            }
          }
        }
      })
    })

    // 商談事務へのコメント
    const oppSql = `SELECT CommentForSalesOffice__c FROM Opportunity WHERE Id = '0060w00000BmPa2AAF'`
    const oppRes = await this.sfdcService.query<Opportunity>(oppSql)
    if (!oppRes || oppRes.totalSize === 0) return
    const opp = oppRes.records[0]
    sheet.getRows(11, 3)?.forEach(row => {
      row.eachCell((cell, colNumber) => {
        this.setCellVal(cell, opp)
      })
    })
  }

  /**
   * SFDCレコードのAPI名により、cell の値を設定する
   * @param cell
   * @param record
   * @param filedApiName
   */
  setCellVal(cell: ExcelJS.Cell, record: any) {
    if (!cell || !cell.value) return
    const valStr = cell.value.toString()
    const regStr = valStr.match("{.*}")
    if (regStr) {
      const filedApiName = regStr[0].substring(1, regStr[0].length - 1)
      if (record[filedApiName]) {
        // 需要家値を設定する
        cell.value = record[filedApiName]
      }
    }
  }

  /**
   * 調査票(新)シートの内容を設定する
   * @param sheet
   */
  private async setQuestionnaireSheet(sheet: ExcelJS.Worksheet) {
    // 需要場所のID
    const optRowMap = new Map<string, CellInfo>()
    sheet.getColumn("C").eachCell((cell, rowNumber) => {
      // 26行目からIdを整理する
      if (rowNumber > 26 && cell.value) {
        const idStr = cell.value.toString()
        const sfdcId = idStr.substring(1, idStr.length - 1)
        optRowMap.set(sfdcId, {
          rowNum: rowNumber,
          sfdcId: sfdcId
        })
      }
    })

    // 需要場所検索
    const sql = `SELECT ${SfdcService.getQueryFileds(
      QuestionnaireService.LocationFiledList
    )} FROM Account WHERE Id IN ${SfdcService.getInSql(optRowMap.keys())}`
    const res = await this.sfdcService.query<Account>(sql)
    if (!res || res.totalSize === 0) return
    const recordMap = new Map<string, Account>()
    for (const acc of res.records) {
      recordMap.set(acc.Id, acc)
    }

    // 内容を出力する
    for (const optRow of optRowMap.values()) {
      sheet.getRow(optRow.rowNum).eachCell((cell, colNumber) => {
        // 開始列 - 施設名
        if (colNumber === 3) {
          const obj = recordMap.get(optRow.sfdcId)
          const filedApiName = "FacilityName__c"
          if (obj) sheet.getCell(optRow.rowNum, colNumber).value = obj[filedApiName] ? obj[filedApiName] : null
        }
        if (colNumber > 3 && cell && cell.value) {
          const valStr = cell.value.toString()
          const regStr = valStr.match("{.*}")
          if (regStr) {
            const filedApiName = regStr[0].substring(1, regStr[0].length - 1)
            const obj = recordMap.get(optRow.sfdcId)
            // 値を設定する
            if (obj) sheet.getCell(optRow.rowNum, colNumber).value = obj[filedApiName] ? obj[filedApiName] : null
          }
        }
      })
    }
  }

  /**
   * シート名により、WorkSheetを取得する
   * @param sheetName シート名
   * @returns
   */
  getWorkSheetByName(sheetName: string): Worksheet | undefined {
    return this.sheetLsit.find(sheet => {
      return sheet.name === sheetName
    })
  }

  /**
   * フォーマットCellにより、同じ型のCellを取得する
   * @param formatCells
   */
  getCellsByFormat(workSheet: ExcelJS.Worksheet, formatCells: ExcelJS.Cell[], PageInfo?: any) {
    if (!PageInfo) {
      PageInfo = this.PageInfo
    }
    const cells: Cell[] = []
    for (let index = 0; index < formatCells.length; index++) {
      const formatCell = formatCells[index]
      const targetCell = workSheet.getCell(PageInfo.rowIndex, PageInfo.colIndex + index)
      targetCell.style = formatCell.style
      cells.push(targetCell)
    }
    return cells
  }
}
