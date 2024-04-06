# Export the data to Excel
$excel = $sortedCombinedUsers | Export-Excel -Path $combinedUsersExcelPath `
    -ClearSheet `
    -AutoSize `
    -AutoFilter `
    -FreezeTopRow `
    -WorksheetName "CombinedUsers" `
    -TableStyle "Medium9" `
    -Title "Combined Users Report" `
    -TitleTop `
    -PassThru

# Add custom styling to the title
$ws = $excel.Workbook.Worksheets["CombinedUsers"]
$titleRow = $ws.Dimension.Start.Row
$titleCell = $ws.Cells[$titleRow, 1]

# Set the horizontal alignment to center
$titleCell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center

# Set the background color to a light blue
$titleCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$titleCell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightBlue)

# Set the font to a more formal style and increase the size
$fontStyle = [System.Drawing.FontStyle]::Bold
$font = New-Object System.Drawing.Font("Calibri", 14, $fontStyle)
$titleCell.Style.Font.SetFromFont($font)

# Save and close the Excel package
$excel.Save()
$excel.Dispose()
