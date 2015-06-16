require_relative './border_creator'

module AxlsxStyler
  module Axlsx
    module Worksheet
      # Example:
      #   add_style 'A1:B5', b: true, sz: 14
      def add_style(cell_ref, style)
        item = self[cell_ref]
        cells = item.is_a?(Array) ? item : [item]
        cells.each { |cell| cell.add_style(style) }
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

      # Example
        # bold = { b: true }
        # large_text = { sz: 30 }
        # add_multiple_styles 'B2:F8', bold, large_text
      def add_multiple_styles(cell_ref, *styles)
        styles.each { |style| add_style(cell_ref, style) }
      end
    end
  end
end
