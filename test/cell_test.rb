require 'test_helper'

class CellTest < MiniTest::Test
  def test_can_add_style
    p = Axlsx::Package.new
    wb = p.workbook
    sheet = wb.add_worksheet
    row = sheet.add_row ["x", "y"]
    cell = row.cells.first

    cell.add_style b: true
    assert_equal({ b: true }, cell.raw_style)
  end
end
