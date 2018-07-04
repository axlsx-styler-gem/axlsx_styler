require 'test_helper'

class GeneralTest < MiniTest::Test

  def setup
    @axlsx = Axlsx::Package.new
    @workbook = @axlsx.workbook
  end

  def test_adding_styled_cells
    p = Axlsx::Package.new
    wb = p.workbook
    wb.add_styled_cell 'Cell 1'
    wb.add_styled_cell 'Cell 2'
    assert_equal ['Cell 1', 'Cell 2'].to_set, wb.styled_cells
    assert_nil wb.styles_applied
  end

  def test_can_add_style
    p = Axlsx::Package.new
    wb = p.workbook
    sheet = wb.add_worksheet
    row = sheet.add_row ["x", "y"]
    cell = row.cells.first

    cell.add_style b: true
    assert_equal({b: true}, cell.raw_style)
  end

  def test_table_with_num_fmt
    filename = 'num_fmt_test'
    t = Time.now
    day = 24 * 60 * 60
    @workbook.add_worksheet do |sheet|
      sheet.add_row %w(Date Count Percent)
      sheet.add_row [t - 2 * day, 2, 2.to_f.round(2) / 11]
      sheet.add_row [t - 1 * day, 3, 3.to_f.round(2) / 11]
      sheet.add_row [t,           6, 6.to_f.round(2) / 11]

      sheet.add_style 'A1:B1', b: true
      sheet.add_style 'A2:A4', format_code: 'YYYY-MM-DD hh:mm:ss'
    end
    serialize(filename)
    assert_equal 2, @workbook.styles.style_index.count
    assert_equal 2 + 2, @workbook.styles.style_index.keys.max
    assert_equal 5, @workbook.styled_cells.count
  end

  def test_duplicate_styles
    filename = 'duplicate_styles'
    @workbook.add_worksheet do |sheet|
      sheet.add_row %w(Index City)
      sheet.add_row [1, 'Ottawa']
      sheet.add_row [2, 'Boston']

      sheet.add_border 'A1:B1', [:bottom]
      sheet.add_border 'A1:B1', [:bottom]
      sheet.add_style 'A1:A3', b: true
      sheet.add_style 'A1:A3', b: true
    end
    serialize(filename)
    assert_equal 4, @workbook.styled_cells.count
    assert_equal 3, @workbook.styles.style_index.count
  end

  def test_multiple_named_styles
    filename = 'multiple_named_styles'
    bold_blue = { b: true, fg_color: '0000FF' }
    large = { sz: 16 }
    red = { fg_color: 'FF0000' }
    @workbook.add_worksheet do |sheet|
      sheet.add_row %w(Index City)
      sheet.add_row [1, 'Ottawa']
      sheet.add_row [2, 'Boston']

      sheet.add_style 'A1:B1', bold_blue, large
      sheet.add_style 'A1:A3', red
    end
    serialize(filename)
    assert_equal 4, @workbook.styled_cells.count
    assert_equal 3, @workbook.styles.style_index.count
  end

end
