require 'set'
require 'active_support/core_ext/hash/deep_merge'

module AxlsxStyler
  module Workbook
    # An array that holds all cells with styles
    attr_accessor :styled_cells

    # Checks if styles are indexed to make it work for pre 0.1.5 version
    # users that still explicitly call @workbook.apply_styles
    attr_accessor :styles_applied

    def add_styled_cell(cell)
      self.styled_cells ||= Set.new
      self.styled_cells << cell
    end

    def apply_styles
      return unless styled_cells

      styled_cells.each do |cell|
        current_style = styles.style_index[cell.style]

        if current_style
          new_style = current_style.deep_merge(cell.raw_style)
        else
          new_style = cell.raw_style
        end

        cell.style = styles.add_style(new_style)
      end

      self.styles_applied = true
    end

  end
end

Axlsx::Workbook.send(:include, AxlsxStyler::Workbook)
