require_relative './border_creator'

module AxlsxStyler
  module Axlsx
    module Worksheet
      # Example to add single style
      #   add_style 'A1:B5', b: true, sz: 14

      # Example to add multiple styles
      #   bold = { b: true }
      #   large_text = { sz: 30 }
      #   add_style 'B2:F8', bold, large_text
      def add_style(cell_ref, *styles)
        item = self[cell_ref]
        cells = item.is_a?(Array) ? item : [item]
        cells.each do |cell|
          styles.each do |style|
            cell.add_style(style)
          end
        end
      end

      # Examples:
      #   add_border 'B2:F8', [:left, :top]
      #   add_border 'C2:G10', [:top]
      #   add_border 'C2:G10'
      # @TODO: allow to pass in custom border style
      def add_border(cell_ref, edges = :all)
        cells = self[cell_ref]
        BorderCreator.new(self, cells, edges).draw
      end
    end
  end
end
