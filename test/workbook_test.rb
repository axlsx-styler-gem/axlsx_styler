class WorkbookTest < MiniTest::Test
  def test_adding_styled_cells
    p = Axlsx::Package.new
    wb = p.workbook
    wb.add_styled_cell 'Cell 1'
    wb.add_styled_cell 'Cell 2'
    assert_equal ['Cell 1', 'Cell 2'].to_set, wb.styled_cells
    assert_nil wb.styles_applied
  end
end
