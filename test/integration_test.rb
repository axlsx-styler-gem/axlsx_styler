require 'test_helper'

class IntegrationTest < MiniTest::Test
  # def test_table_with_borders
  #   axlsx = Axlsx::Package.new
  #   workbook = axlsx.workbook
  #   workbook.add_worksheet do |sheet|
  #     sheet.add_row
  #     sheet.add_row ['', 'Product', 'Category',  'Price']
  #     sheet.add_row ['', 'Butter', 'Dairy',      4.99]
  #     sheet.add_row ['', 'Bread', 'Baked Goods', 3.45]
  #     sheet.add_row ['', 'Broccoli', 'Produce',  2.99]
  #     sheet.add_row ['', 'Pizza', 'Frozen Foods',  4.99]
  #     sheet.column_widths 5, 20, 20, 20

  #     # using AxlsxStyler DSL
  #     sheet['B2:D2'].add_style b: true
  #     sheet['B2:B6'].add_style b: true
  #     sheet['B2:D2'].add_style bg_color: '95AFBA'
  #     sheet['B3:D6'].add_style bg_color: 'E2F89C'
  #     sheet['D3:D6'].add_style alignment: { horizontal: :left }
  #     sheet['B2:D6'].add_border
  #     sheet['B3:D3'].add_border [:top]
  #   end
  #   workbook.apply_styles
  #   axlsx.serialize File.expand_path(
  #     '../../tmp/borders_test.xlsx',
  #     __FILE__
  #   )
  #   assert_equal 12, workbook.style_index.count
  #   assert_equal 12 + 2, workbook.style_index.keys.max
  # end

  def test_table_with_num_fmt
    axlsx = Axlsx::Package.new
    workbook = axlsx.workbook
    t = Time.now
    day = 24 * 60 * 60
    workbook.add_worksheet do |sheet|
      sheet.add_row %w(Date Count Percent)
      sheet.add_row [t - 2 * day, 2, 2.to_f / 11]
      sheet.add_row [t - 1 * day, 3, 3.to_f / 11]
      sheet.add_row [t,           6, 6.to_f / 11]

      sheet.add_style 'A1:B1', b: true
      sheet.add_style 'A2:A4', format_code: 'YYYY-MM-DD hh:mm:ss'
    end
    workbook.apply_styles
    assert_equal 2, workbook.style_index.count
    axlsx.serialize File.expand_path(
      '../../tmp/num_fmt_test.xlsx',
      __FILE__
    )
  end

  # def test_duplicate_styles
  #   axlsx = Axlsx::Package.new
  #   workbook = axlsx.workbook
  #   workbook.add_worksheet do |sheet|
  #     sheet.add_row %w(Index City)
  #     sheet.add_row [1, 'Ottawa']
  #     sheet.add_row [2, 'Boston']

  #     sheet['A1:B1'].add_border [:bottom]
  #     sheet['A1:B1'].add_border [:bottom]
  #     sheet['A1:A3'].add_style b: true
  #     sheet['A1:A3'].add_style b: true
  #   end
  #   workbook.apply_styles
  #   assert_equal 4, workbook.styled_cells.count
  #   assert_equal 3, workbook.style_index.count
  #   axlsx.serialize File.expand_path(
  #     '../../tmp/duplicate_styles.xlsx',
  #     __FILE__
  #   )
  # end
end
