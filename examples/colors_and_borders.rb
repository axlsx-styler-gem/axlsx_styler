require_relative '../lib/axlsx_styler'

axlsx = Axlsx::Package.new
workbook = axlsx.workbook

workbook.add_worksheet do |sheet|
  sheet.add_row
  sheet.add_row ["", "Product", "Category", "Price"]
  sheet.add_row ["", "Butter", "Dairy", 4.99]
  sheet.add_row ["", "Bread", "Baked Goods", 3.45]
  sheet.add_row ["", "Broccoli", "Produce", 2.99]

  sheet["B2:D2"].add_style(b: true)
  sheet["B2:D5"].add_style(bg_color: "E2D3EB")
  sheet["B2:D5"].add_border
  sheet["B3:D3"].add_border([:top])
end

workbook.apply_styles

axlsx.serialize "grocery.xlsx"
