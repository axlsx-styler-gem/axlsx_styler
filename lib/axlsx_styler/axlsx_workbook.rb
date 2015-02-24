module AxlsxStyler
  module Axlsx
    module Workbook
      attr_accessor :styled_cells

      def add_styled_cell(cell)
        self.styled_cells ||= Set.new
        self.styled_cells << cell
      end

      def apply_styles
        return unless styled_cells
        styled_cells.each do |cell|
          cell.style = styles.add_style(cell.raw_style)
        end
      end
    end
  end
end
