require 'axlsx_styler'

axlsx = Axlsx::Package.new
workbook = axlsx.workbook
workbook.add_worksheet do |sheet|
  sheet.add_row
  sheet.add_row ['', 'Product', 'Category',  'Price']
  sheet.add_row ['', 'Butter', 'Dairy',      4.99]
  sheet.add_row ['', 'Bread', 'Baked Goods', 3.45]
  sheet.add_row ['', 'Broccoli', 'Produce',  2.99]
  sheet.column_widths 5, 20, 20, 20

  # using AxlsxStyler DSL
  sheet.add_style 'B2:D2', b: true
  sheet.add_style 'B2:B5', b: true
  sheet.add_style 'B2:D2', bg_color: '95AFBA'
  sheet.add_style 'B3:D5', bg_color: 'E2F89C'
  sheet.add_style 'D3:D5', alignment: { horizontal: :left }
  sheet.add_style ['C3:C4', 'D3:D4'], fg_color: '00FF00'
  sheet.add_border 'B2:D5'
  sheet.add_border 'B3:D3', { edges: [:top], style: :thick }
  sheet.add_border ['C3:C4', 'D3:D4']
end
axlsx.serialize File.expand_path('../../tmp/grocery.xlsx', __FILE__)
