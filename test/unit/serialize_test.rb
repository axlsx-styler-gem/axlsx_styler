require 'test_helper'

class SerializeTest < MiniTest::Test

  def setup
    @axlsx = Axlsx::Package.new
    @workbook = @axlsx.workbook
  end

  def test_works_without_apply_styles_serialize
    filename = 'without_apply_styles_serialize'
    assert_nil @workbook.styles_applied
    @workbook.add_worksheet do |sheet|
      sheet.add_row ['A1', 'B1']
      sheet.add_style 'A1:B1', b: true
    end
    serialize(filename)
    assert_equal 1, @workbook.styles.style_index.count
  end

  # Backwards compatibility with pre 0.1.5 (serialize)
  def test_works_with_apply_styles_serialize
    filename = 'with_apply_styles_serialize'
    assert_nil @workbook.styles_applied
    @workbook.add_worksheet do |sheet|
      sheet.add_row ['A1', 'B1']
      sheet.add_style 'A1:B1', b: true
    end
    @workbook.apply_styles # important for backwards compatibility
    assert_equal 1, @workbook.styles.style_index.count
    serialize(filename)
  end

end
