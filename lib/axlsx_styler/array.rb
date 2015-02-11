class Array
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
    selected_edges(edges).each{ |edge| add_border_at(edge) }
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
    # example range "B2:D5"
    {
      top:     "#{first}:#{last[0]}#{first[1]}", # "B2:D2"
      right:   "#{last[0]}#{first[1]}:#{last}",  # "D2:D5"
      bottom:  "#{first[0]}#{last[1]}:#{last}",  # "B5:D5"
      left:    "#{first}:#{first[0]}#{last[1]}"  # "B2:B5"
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
