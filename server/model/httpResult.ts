export interface HttpResult {
  status: number // 0: 成功 1:警告 2:error
  data: any
  msg: string
}
