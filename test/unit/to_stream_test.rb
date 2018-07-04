require 'test_helper'

class ToStreamTest < MiniTest::Test

  def setup
    @axlsx = Axlsx::Package.new
    @workbook = @axlsx.workbook
  end

  def test_works_without_apply_styles_to_stream
    filename = 'without_apply_styles_to_stream'
    assert_nil @workbook.styles_applied
    @workbook.add_worksheet do |sheet|
      sheet.add_row ['A1', 'B1']
      sheet.add_style 'A1:B1', b: true
    end
    to_stream(filename)
    assert_equal 1, @workbook.styles.style_index.count
  end

  # Backwards compatibility with pre 0.1.5 (to_stream)
  def test_works_with_apply_styles_to_stream
    filename = 'with_apply_styles_to_stream'
    assert_nil @workbook.styles_applied
    @workbook.add_worksheet do |sheet|
      sheet.add_row ['A1', 'B1']
      sheet.add_style 'A1:B1', b: true
    end
    @workbook.apply_styles # important for backwards compatibility
    assert_equal 1, @workbook.styles.style_index.count
    to_stream(filename)
  end

end
