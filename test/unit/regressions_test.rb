require 'test_helper'

class RegressionsTest < MiniTest::Test

  def setup
    @axlsx = Axlsx::Package.new
    @workbook = @axlsx.workbook
  end

  def test_dxf_cell
    @workbook.add_worksheet do |sheet|
      sheet.add_row (1..2).to_a
      sheet.add_style "A1:A1", { bg_color: "AA0000" }

      sheet.add_row (1..2).to_a
      sheet.add_style "B1:B1", { bg_color: "CC0000" }

      sheet.add_row (1..2).to_a
      sheet.add_style "A3:B3", { bg_color: "00FF00" }

      highlight = @workbook.styles.add_style(bg_color: "0000FF", type: :dxf)

      sheet.add_conditional_formatting(
        "A2:B2",
        {
          type: :cellIs,
          operator: :greaterThan,
          formula: "1",
          dxfId: highlight,
          priority: 1
        }
      )
    end

    serialize("test_dxf_cell")

    #puts @workbook.styles.dxfs.map{|x| x.to_xml_string}
    assert_equal @workbook.styles.dxfs.count, 1

    #puts @workbook.styles.cellXfs.map{|x| x.to_xml_string}
    assert_equal @workbook.styles.cellXfs.count, 6
  end

  def test_custom_default_font_and_size
    @workbook.styles.fonts[0].name = 'Pontiac' ### TODO, is this a valid font name in all environments
    @workbook.styles.fonts[0].sz = 12

    @workbook.add_worksheet do |sheet|
      sheet.add_row [1,2,3]
      sheet.add_style "A1:C1", { color: "FFFFFF" }
    end

    @workbook.apply_styles

    assert_equal 1, @workbook.styles.style_index.size
    
    assert_equal(
      {
        type: :xf, 
        name: "Pontiac", 
        sz: 12, 
        family: 1, 
        color: "FFFFFF",
      }, 
      @workbook.styles.style_index.values.first
    )

    parsed_sheet = Nokogiri::XML(@workbook.worksheets.first.to_xml_string) 
    parsed_styles = Nokogiri::XML(@workbook.worksheets.first.to_xml_string) 

    ### TODO want to examine xml contents to ensure font working
    ### Could possibly just do manual testing if required but should be automatable
    
    # puts parsed_sheet.to_xml_string
    # puts parsed_styles.to_xml_string

    binding.pry
  end

end
