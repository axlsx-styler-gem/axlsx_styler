require 'test_helper'

class IntegrationTest < MiniTest::Test
  def setup
    @axlsx = Axlsx::Package.new
    @workbook = @axlsx.workbook
  end

  # Save to a file using Axlsx::Package#serialize
  def serialize(filename)
    @axlsx.serialize File.expand_path("../../tmp/#{filename}.xlsx", __FILE__)
    assert_equal true, @workbook.styles_applied
  end

  # Save to a file by getting contents from stream
  def to_stream(filename)
    File.open(File.expand_path("../../tmp/#{filename}.xlsx", __FILE__), 'w') do |f|
      f.write @axlsx.to_stream.read
    end
  end

  # New functionality as of 0.1.5 (serialize)
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

  # New functionality as of 0.1.5 (to_stream)
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
