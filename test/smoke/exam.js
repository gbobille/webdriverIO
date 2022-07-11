// @flow
import { assert } from 'chai'
// @flow
import ExcelJs from 'exceljs'

const workbook = new ExcelJs.Workbook()
let workbookName
let worksheet
let text

export default {
  /**
   * Function to get cell data from given excel file and sheet and cell.
   * @param filename - Excel file name
   * @param sheetid - Excel file's worksheet id
   * @param cellId - Excel file's worksheet's cell if
   * @param retries - by default function run once but can be increase
   * @returns {Promise<*>} - return specified cell data (value)
   */
  async getCellData (filename: string, sheetid: number, cellId: string, retries: number = 1): Promise<any> {
    try {
      // $FlowIssue getting xlsx options as readFile method is a-ok
      workbookName = await workbook.xlsx.readFile(filename)
      worksheet = await workbookName.getWorksheet(sheetid)
      text = await worksheet.getCell(cellId).value
      return text
    } catch (err) {
      if (retries === 0) {
        throw new Error(`Cell ${cellId.toString()} value is not found`)
      }
      return this.getCellData(filename, sheetid, cellId, retries - 1)
    }
  }
}

let baseUrl = process.env.SELENIUM_SERVER_URL === 'qa' ? `https://${process.env.SELENIUM_SERVER_URL}.google.com/` : `https://${process.env.SELENIUM_SERVER_URL}.google.com/`
const datsearch = async (row: number = 2, sheet: number = 1) => ExcelUtil.getCellData(excelFileName, sheet, `A${await row}`)
const SEACH_FIELD = '//input[@name="q"]'
const SEACH_BUTTON = '//input[@name="btnK"]'
describe('Search for a keyword', function () {
  before(async function () {
    await browser.url(baseUrl)
  })

  it('Keyword Searcch', async function () {
    await browser.sendKeys(SEACH_FIELD, datsearch)
    await browser.touchClick(SEACH_BUTTON)
    assert.hasAnyKeys(datsearch, 'Something went wrong happened in the page')
  })
})
