import { Account } from "../model/account"

export class Tools {
  /**
   * テキスト to Date
   * @param dateStr 時間フォーマットのテキスト 例:2020-01-15   or   2022/02/15
   * @returns Date
   */
  public static strToDate(dateStr: string | Date | null | undefined): Date | undefined {
    if (!dateStr) {
      return undefined
    }
    if (dateStr instanceof Date) {
      return dateStr
    }
    const exportDate = new Date(dateStr.replace(/-/g, "/"))
    const isInvalidDate = (date: Date) => Number.isNaN(date.getTime())
    if (isInvalidDate(exportDate)) {
      throw new Error(dateStr + "が有効な日付ではあまりせん")
    }
    return exportDate
  }

  /**
   * 2つ日付間の日数を求める
   * @param date1
   * @param date2
   * @returns
   */
  public static getDays(date1: Date, date2: Date): number {
    return (date2.getTime() - date1.getTime()) / 86400000 + 1
  }

  /**
   * 曜日取得する
   * @param date
   * @returns
   */
  public static getWeekNameForJP(date: Date) {
    const dayOfWeek = date.getDay()
    return ["日", "月", "火", "水", "木", "金", "土"][dayOfWeek]
  }
}
