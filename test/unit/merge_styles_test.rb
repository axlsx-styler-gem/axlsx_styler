require 'test_helper'

class MergeStylesTest < MiniTest::Test

  def setup
    @axlsx = Axlsx::Package.new
    @workbook = @axlsx.workbook
  end

  def test_merge_styles_1
    filename = 'merge_styles_1'
    bold = @workbook.styles.add_style b: true

    @workbook.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row ['', '1', '2', '3'], style: [nil, bold]
      sheet.add_row ['', '4', '5', '6'], style: bold
      sheet.add_row ['', '7', '8', '9']
      sheet.add_style 'B2:D4', b: true
      sheet.add_border 'B2:D4', { style: :thin, color: '000000' }
    end
    @workbook.apply_styles
    assert_equal 9, @workbook.styles.style_index.count
    serialize(filename)
  end

  def test_merge_styles_2
    filename = 'merge_styles_2'
    bold = @workbook.styles.add_style b: true

    @workbook.add_worksheet do |sheet|
      sheet.add_row ['A1', 'B1'], style: [nil, bold]
      sheet.add_row ['A2', 'B2'], style: bold
      sheet.add_row ['A3', 'B3']
      sheet.add_style 'A1:A2', i: true
    end
    @workbook.apply_styles
    assert_equal 3, @workbook.styles.style_index.count
    serialize(filename)
  end

  def test_merge_styles_3
    filename = 'merge_styles_3'
    bold = @workbook.styles.add_style b: true

    @workbook.add_worksheet do |sheet|
      sheet.add_row ['A1', 'B1'], style: [nil, bold]
      sheet.add_row ['A2', 'B2']
      sheet.add_style 'B1:B2', bg_color: 'FF0000'
    end
    @workbook.apply_styles
    assert_equal 3, @workbook.styles.style_index.count
    serialize(filename)
  end

end
