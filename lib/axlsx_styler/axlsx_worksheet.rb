require_relative './border_creator'

module AxlsxStyler
  module Worksheet
    # Example to add a single style:
    #   add_style 'A1:B5', b: false
    #   add_style ['B1', 'B3'], b: false
    #   add_style ['D3:D4', 'F2:F6'], b: false 
    #
    # Example to add multiple styles:
    #   bold = { b: true }
    #   large_text = { sz: 30 }
    #   add_style 'B2:F8', bold, large_text
    def add_style(cell_refs, *styles)
      if !cell_refs.is_a?(Array)
        cell_refs = [cell_refs]
      end

      cell_refs.each do |cell_ref|
        item = self[cell_ref]

        cells = item.is_a?(Array) ? item : [item]

        cells.each do |cell|
          styles.each do |style|
            cell.add_style(style)
          end
        end
      end
    end

    # Examples:
    #   add_border 'B2:F8', [:left, :top], :medium, '00330f'
    #   add_border 'B2:F8', [:left, :top], :medium
    #   add_border 'C2:G10', [:top]
    #   add_border 'C2:G10'
    #   add_border 'B2:D5', { style: :thick, color: '00330f', edges: [:left, :right] }
    #   add_border ['D3:D4', 'F2:F6'], [:left]
    def add_border(cell_refs, args = :all)
      if !cell_refs.is_a?(Array)
        cell_refs = [cell_refs]
      end

      cell_refs.each do |cell_ref|
        cells = self[cell_ref]
        BorderCreator.new(self, cells, args).draw
      end
    end
  end
end

Axlsx::Worksheet.send(:include, AxlsxStyler::Worksheet)
