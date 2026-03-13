import ExcelJS from 'exceljs'

const filePath = 'public/8300 Meadowbrook Ln_WGCPPS1550333380_Final ECM Calculations.xlsx'
const wb = new ExcelJS.Workbook()
await wb.xlsx.readFile(filePath)

const sourceWs = wb.getWorksheet('Input Sheet')
const targetWs = wb.getWorksheet("Input Sheet -W'Gas")

if (!sourceWs || !targetWs) {
  console.error('Worksheets not found')
  process.exit(1)
}

console.log(`Copying content from ${sourceWs.name} to ${targetWs.name}`)

// Copy Columns
for (let i = 1; i <= sourceWs.columnCount; i++) {
  const col = sourceWs.getColumn(i)
  const targetCol = targetWs.getColumn(i)
  targetCol.width = col.width
  targetCol.style = col.style
}

// Copy Rows and Cells
sourceWs.eachRow({ includeEmpty: true }, (row, rowNumber) => {
  const targetRow = targetWs.getRow(rowNumber)
  targetRow.height = row.height

  row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const targetCell = targetRow.getCell(colNumber)
    targetCell.value = cell.value
    targetCell.style = cell.style
    targetCell.numFmt = cell.numFmt
    targetCell.dataValidation = cell.dataValidation
  })
})

// Copy Merges
if (sourceWs.model.merges) {
  sourceWs.model.merges.forEach((merge) => {
    try {
      targetWs.mergeCells(merge)
    } catch (e) {
      // ignore
    }
  })
}

// Copy Tables
if (sourceWs.tables) {
  for (const name of Object.keys(sourceWs.tables)) {
    const table = sourceWs.tables[name]
    const newTableName = `${name}_Project`

    const columns = table.columns.map((c) => ({
      name: c.name,
      totalsRowLabel: c.totalsRowLabel,
      totalsRowFunction: c.totalsRowFunction,
      totalsRowFormula: c.totalsRowFormula,
      filterButton: c.filterButton,
    }))

    try {
      targetWs.addTable({
        name: newTableName,
        ref: table.ref,
        headerRow: table.headerRow,
        totalsRow: table.totalsRow,
        style: table.style,
        columns: columns,
        rows: table.rows,
      })
      console.log(`Copied table ${name} as ${newTableName}`)
    } catch (e) {
      console.error(`Failed to copy table ${name}:`, e.message)
    }
  }
}

await wb.xlsx.writeFile(filePath)
console.log('Done')
