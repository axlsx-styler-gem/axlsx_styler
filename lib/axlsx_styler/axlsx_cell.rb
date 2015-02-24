require 'active_support/core_ext/hash/deep_merge'

module AxlsxStyler
  module Axlsx
    module Cell
      attr_accessor :raw_style

      def workbook
        row.worksheet.workbook
      end

      def add_style(style)
        self.raw_style ||= {}
        add_to_raw_style(style)
        workbook.add_styled_cell self
      end

      def add_to_raw_style(style)
        # using deep_merge from active_support:
        # with regular Hash#merge adding borders fails miserably
        new_style = raw_style.deep_merge style
        if with_border?(raw_style) && with_border?(style)
          border_at = raw_style[:border][:edges] + style[:border][:edges]
          new_style[:border][:edges] = border_at
        elsif with_border?(style)
          new_style[:border] = style[:border]
        end
        self.raw_style = new_style
      end

      def with_border?(style)
        !style[:border].nil?
      end
    end
  end
end
