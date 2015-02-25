module AxlsxStyler
  # @TODO
  # Move the functionality from Array to Worksheet. New DSL
  #   sheet.add_style 'A1:B5', b: true
  # instead of
  #   sheet['A1:B5'].add_style, b: true
  module Array
    def add_style(style)
      validate_cells
      each do |cell|
        cell.add_style(style)
      end
    end

    # - edges is either :all or an an array with elements
    #   in [:top, :bottom, :bottom, :left]
    # - couldn't do variable border thickness around the same cell;
    #   hardcode style to :thin for now
    def add_border(edges = :all)
      validate_cells
      selected_edges(edges).each { |edge| add_border_at(edge) }
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

    def validate_cells
      valid = map{ |e| e.is_a? Axlsx::Cell }.uniq == [ true ]
      raise "Not a range of cells" unless valid
    end

    def border_cells
      @border_cells ||= get_border_cells
    end

    def get_border_cells
      first = self.first.r
      last  = self.last.r

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

    def worksheet
      @worksheet ||= self.first.row.worksheet
    end

    def add_border_at(position)
      style = {
        border: {
          style: :thin, color: "000000", edges: [position.to_sym]
        }
      }
      worksheet[border_cells[position.to_sym]].add_style(style)
    end
  end
end
