require 'test_helper'

class IntegrationTest < MiniTest::Test
  # taken from /examples
  def test_table_with_borders
    axlsx = Axlsx::Package.new
    workbook = axlsx.workbook
    workbook.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row ['', 'Product', 'Category',  'Price']
      sheet.add_row ['', 'Butter', 'Dairy',      4.99]
      sheet.add_row ['', 'Bread', 'Baked Goods', 3.45]
      sheet.add_row ['', 'Broccoli', 'Produce',  2.99]
      sheet.add_row ['', 'Pizza', 'Frozen Foods',  4.99]
      sheet.column_widths 5, 20, 20, 20

      # using AxlsxStyler DSL
      sheet['B2:D2'].add_style b: true
      sheet['B2:B6'].add_style b: true
      sheet['B2:D2'].add_style bg_color: '95AFBA'
      sheet['B3:D6'].add_style bg_color: 'E2F89C'
      sheet['D3:D6'].add_style alignment: { horizontal: :left }
      sheet['B2:D6'].add_border
      sheet['B3:D3'].add_border [:top]
    end
    workbook.apply_styles
    axlsx.serialize File.expand_path(
      '../../tmp/table_with_borders_test.xlsx',
      __FILE__
    )
    assert_equal 12, workbook.style_index.count
  end
end
