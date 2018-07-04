require 'test_helper'

class BordersTest < MiniTest::Test

  def setup
    @axlsx = Axlsx::Package.new
    @workbook = @axlsx.workbook
  end

  def test_table_with_borders
    filename = 'borders_test'
    @workbook.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row ['', 'Product', 'Category',  'Price']
      sheet.add_row ['', 'Butter', 'Dairy',      4.99]
      sheet.add_row ['', 'Bread', 'Baked Goods', 3.45]
      sheet.add_row ['', 'Broccoli', 'Produce',  2.99]
      sheet.add_row ['', 'Pizza', 'Frozen Foods',  4.99]
      sheet.column_widths 5, 20, 20, 20

      sheet.add_style 'B2:D2', b: true
      sheet.add_style 'B2:B6', b: true
      sheet.add_style 'B2:D2', bg_color: '95AFBA'
      sheet.add_style 'B3:D6', bg_color: 'E2F89C'
      sheet.add_style 'D3:D6', alignment: { horizontal: :left }
      sheet.add_border 'B2:D6'
      sheet.add_border 'B3:D3', [:top]
      sheet.add_border 'B3:D3', edges: [:bottom], style: :medium
      sheet.add_border 'B3:D3', edges: [:bottom], style: :medium, color: '32f332'
    end
    serialize(filename)
    assert_equal 12, @workbook.styles.style_index.count
    assert_equal 12 + 2, @workbook.styles.style_index.keys.max
  end

  def test_duplicate_borders
    filename = 'duplicate_borders_test'
    @workbook.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row ['', 'B2', 'C2', 'D2']
      sheet.add_row ['', 'B3', 'C3', 'D3']
      sheet.add_row ['', 'B4', 'C4', 'D4']

      sheet.add_border 'B2:D4'
      sheet.add_border 'B2:D4'
    end
    serialize(filename)
    assert_equal 8, @workbook.styles.style_index.count
    assert_equal 8, @workbook.styled_cells.count
  end

  def test_multiple_style_borders_on_same_cells
    filename = 'multiple_style_borders'
    @workbook.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row ['', 'B2', 'C2', 'D2']
      sheet.add_row ['', 'B3', 'C3', 'D3']

      sheet.add_border 'B2:D3', :all
      sheet.add_border 'B2:D2', edges: [:bottom], style: :thick, color: 'ff0000'
    end
    serialize(filename)
    assert_equal 6, @workbook.styles.style_index.count
    assert_equal 6, @workbook.styled_cells.count

    b2_cell_style = {
      border: {
        style: :thick,
        color: 'ff0000',
        edges: [:bottom, :left, :top]
      },
      type: :xf,
      name: 'Arial',
      sz: 11,
      family: 1
    }
    assert_equal b2_cell_style, @workbook.styles.style_index.values.find{|x| x == b2_cell_style}

    d3_cell_style = {
      border: {
        style: :thin,
        color: '000000',
        edges: [:bottom, :right]
      },
      type: :xf,
      name: 'Arial',
      sz: 11,
      family: 1
    }
    assert_equal d3_cell_style, @workbook.styles.style_index.values.find{|x| x == d3_cell_style}
  end

  # Overriding borders (part 1)
  def test_mixed_borders_1
    filename = 'mixed_borders_1'
    @workbook.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row ['', '1', '2', '3']
      sheet.add_row ['', '4', '5', '6']
      sheet.add_row ['', '7', '8', '9']
      sheet.add_style 'B2:D4', border: { style: :thin, color: '000000' }
      sheet.add_border 'C3:D4', style: :medium
    end
    @workbook.apply_styles
    assert_equal 9, @workbook.styled_cells.count
    assert_equal 2, @workbook.styles.style_index.count
    serialize(filename)
  end

  # Overriding borders (part 2)
  def test_mixed_borders
    filename = 'mixed_borders_2'
    @workbook.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row ['', '1', '2', '3']
      sheet.add_row ['', '4', '5', '6']
      sheet.add_row ['', '7', '8', '9']
      sheet.add_border 'B2:D4', style: :medium
      sheet.add_style 'D2:D4', border: { style: :thin, color: '000000' }
    end
    @workbook.apply_styles
    assert_equal 8, @workbook.styled_cells.count
    assert_equal 6, @workbook.styles.style_index.count
    serialize(filename)
  end

end
