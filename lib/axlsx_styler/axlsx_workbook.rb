module AxlsxStyler
  module Axlsx
    module Workbook
      # An array that holds all cells with styles
      attr_accessor :styled_cells

      # An index for cell styles
      #   {
      #     < style_hash > => 1,
      #     < style_hash > => 2,
      #     ...
      #     < style_hash > => K
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
      end

      private

      # Check if style code
      def set_style_index(cell)
        # @TODO fix this hack
        self.style_index ||= {}

        style = style_index[cell.raw_style]
        if style
          cell.style = style
        else
          new_style = styles.add_style(cell.raw_style)
          cell.style = new_style

          # :num_fmt is distinct even though the styles are
          # the same; not sure if it's intended functionality
          cell.raw_style.delete(:num_fmt)

          style_index[cell.raw_style] = new_style
        end
      end
    end
  end
end
