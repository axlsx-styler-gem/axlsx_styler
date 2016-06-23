require 'set'

module AxlsxStyler
  module Axlsx
    module Workbook
      # An array that holds all cells with styles
      attr_accessor :styled_cells

      # Checks if styles are idexed to make it work for pre 0.1.4 version
      # users that still explicitly call @workbook.apply_styles
      attr_accessor :styles_applied

      # An index for cell styles
      #   {
      #     1 => < style_hash >,
      #     2 => < style_hash >,
      #     ...
      #     K => < style_hash >
      #   }
      # where keys are Cell#raw_style and values are styles
      # codes as per Axlsx::Style
      attr_accessor :style_index

      def add_styled_cell(cell)
        self.styled_cells ||= Set.new
        self.styled_cells << cell
      end

      def apply_styles
        return unless styled_cells
        styled_cells.each do |cell|
          set_style_index(cell)
        end
        self.styles_applied = true
      end

      private

      # Check if style code
      def set_style_index(cell)
        self.style_index ||= {}

        index_item = style_index.select { |_, v| v == cell.raw_style }.first
        if index_item
          cell.style = index_item.first
        else
          old_style = cell.raw_style.dup
          new_style = styles.add_style(cell.raw_style)
          cell.style = new_style
          # cell.raw_style.delete(:num_fmt)
          style_index[new_style] = old_style
        end
      end
    end
  end
end
