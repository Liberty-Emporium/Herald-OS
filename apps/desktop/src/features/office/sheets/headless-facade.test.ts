import { describe, expect, it } from 'vitest'
// A window showing a workbook loads the Sheets UI facade; once loaded, its listeners run on every
// Univer instance, including a headless one — which registers none of the UI services they reach
// for (hover and drag managers). See headless.ts.
import '@univerjs/sheets-ui/facade'
import { featureWorkbook } from '../../../../shared/office/xlsx/fixtures.ts'
import { workbookFromXlsx } from '../../../../shared/office/xlsx/read.ts'
import { withHeadlessSheets } from './headless.ts'

describe('a headless workbook after a window has loaded the Sheets UI facade', () => {
  it('runs without the hover and drag managers the headless instance does not register', async () => {
    const direct = await workbookFromXlsx(await featureWorkbook(), { id: 'book', name: 'Book' })
    const { snapshot } = await withHeadlessSheets(direct.workbook, () => undefined)

    expect(snapshot.id).toBe('book')
  })
})
