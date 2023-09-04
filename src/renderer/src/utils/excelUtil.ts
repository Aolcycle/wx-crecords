import * as XLSX from 'xlsx'

/**
 * 导出
 * @param json 数据
 * @param name 文件名
 * @param titleArr 标题
 * @param sheetName 表名
 */
export const exportExcel = (
  json: [],
  name: string,
  titleArr: string[],
  sheetName: string
): void => {
  /* convert state to workbook */
  const data: Array<string[]> = []
  const keyArray: string[] = []
  const getLength = (obj): number => {
    let count = 1
    for (const i in obj) {
      // eslint-disable-next-line no-prototype-builtins
      if (obj.hasOwnProperty(i)) {
        count++
      }
    }
    return count
  }

  if (!Array.isArray(titleArr)) {
    titleArr = []
  }

  for (const key1 in json) {
    // eslint-disable-next-line no-prototype-builtins
    if (json.hasOwnProperty(key1)) {
      const element: object = json[key1]
      const rowDataArray: string[] = []
      for (const key2 in element) {
        // eslint-disable-next-line no-prototype-builtins
        if (element.hasOwnProperty(key2)) {
          const element2 = element[key2]
          rowDataArray.push(element2)
          if (keyArray.length < getLength(element)) {
            keyArray.push(key2)
          }
        }
      }
      data.push(rowDataArray)
    }
  }
  // keyArray为英文字段表头
  data.splice(0, 0, titleArr)
  const ws = XLSX.utils.aoa_to_sheet(data)
  const wb = XLSX.utils.book_new()
  // 此处隐藏英文字段表头
  // var wsrows = [{ hidden: true }];
  // ws['!rows'] = wsrows; // ws - worksheet
  XLSX.utils.book_append_sheet(wb, ws, sheetName)
  /* generate file and send to client */
  XLSX.writeFile(wb, name + '.xlsx')
}
