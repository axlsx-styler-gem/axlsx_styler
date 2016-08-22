require 'set'

module AxlsxStyler
  module Axlsx
    module Workbook
      # An array that holds all cells with styles
      attr_accessor :styled_cells

      # Checks if styles are idexed to make it work for pre 0.1.5 version
      # users that still explicitly call @workbook.apply_styles
      attr_accessor :styles_applied

      def add_styled_cell(cell)
        self.styled_cells ||= Set.new
        self.styled_cells << cell
      end

      def apply_styles
        return unless styled_cells
        styled_cells.each do |cell|
          cell.style = styles.add_style(cell.raw_style)
        end
        self.styles_applied = true
      end

    end
  end
end
