class BorderCreator
  attr_reader :worksheet, :cells, :edges, :width, :color

  def initialize(worksheet, cells, edges, width, color)
    @worksheet = worksheet
    @cells     = cells
    @edges     = edges
    @width     = width
    @color     = color
  end

  def draw
    selected_edges(edges).each { |edge| add_border(edge, width, color) }
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

  def add_border(position, width, color)
    style = {
      border: {
        style: width, color: color, edges: [position.to_sym]
      }
    }
    worksheet.add_style border_cells[position.to_sym], style
    # add_style border_cells(cell_ref)[position.to_sym], style
  end

  def border_cells
    # example range "B2:D5"
    {
      top:     "#{first_cell}:#{last_col}#{first_row}", # "B2:D2"
      right:   "#{last_col}#{first_row}:#{last_cell}",  # "D2:D5"
      bottom:  "#{first_col}#{last_row}:#{last_cell}",  # "B5:D5"
      left:    "#{first_cell}:#{first_col}#{last_row}"  # "B2:B5"
    }
  end

  def first_cell
    @first_cell ||= cells.first.r
  end

  def last_cell
    @last_cell ||= cells.last.r
  end

  def first_row
    @first_row ||= first_cell.scan(/\d+/).first
  end

  def first_col
    @first_col ||= first_cell.scan(/\D+/).first
  end

  def last_row
    @last_row ||= last_cell.scan(/\d+/).first
  end

  def last_col
    @last_col ||= last_cell.scan(/\D+/).first
  end
end
