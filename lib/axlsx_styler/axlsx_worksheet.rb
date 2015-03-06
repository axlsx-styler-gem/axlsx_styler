module AxlsxStyler
  module Axlsx
    module Worksheet
      # Example:
      #   add_style 'A1:B5', b: true, sz: 14
      def add_style(cell_ref, style)
        self[cell_ref].each do |cell|
          cell.add_style(style)
        end
      end

      # Examples:
      #   add_border 'B2:F8', [:left, :top]
      #   add_border 'C2:G10', [:top]
      #   add_border 'C2:G10'
      def add_border(cell_ref, edges = :all)
        selected_edges(edges).each { |edge| add_border_at(cell_ref, edge) }
      end

      private

      def selected_edges(edges)
        all_edges = [:top, :right, :bottom, :left]
        if edges == :all
          all_edges
        elsif edges.is_a?(Array) && edges - all_edges == []
          edges.uniq
        else
          []
        end
      end

      # @TODO: allow for custom border width
      def add_border_at(cell_ref, position)
        style = {
          border: {
            style: :thin, color: '000000', edges: [position.to_sym]
          }
        }
        add_style border_cells(cell_ref)[position.to_sym], style
      end

      def border_cells(cell_ref)
        cells = self[cell_ref]
        first = cells.first.r
        last  = cells.last.r

        first_row = first.scan(/\d+/).first
        last_row  = last.scan(/\d+/).first
        first_col = first.scan(/\D+/).first
        last_col  = last.scan(/\D+/).first

        # example range "B2:D5"
        {
          top:     "#{first}:#{last_col}#{first_row}", # "B2:D2"
          right:   "#{last_col}#{first_row}:#{last}",  # "D2:D5"
          bottom:  "#{first_col}#{last_row}:#{last}",  # "B5:D5"
          left:    "#{first}:#{first_col}#{last_row}"  # "B2:B5"
        }
      end
    end
  end
end
