require 'axlsx'

require 'axlsx_styler/version'
require 'axlsx_styler/axlsx_workbook'
require 'axlsx_styler/axlsx_worksheet'
require 'axlsx_styler/axlsx_cell'

Axlsx::Workbook.send :include, AxlsxStyler::Axlsx::Workbook
Axlsx::Worksheet.send :include, AxlsxStyler::Axlsx::Worksheet
Axlsx::Cell.send :include, AxlsxStyler::Axlsx::Cell

module Axlsx
  class Package
    # Patches the original Axlsx::Package#serialize method so that styles are
    # applied when the workbook is saved
    original_serialize = instance_method(:serialize)
    define_method :serialize do |*args|
      workbook.apply_styles if !workbook.styles_applied
      original_serialize.bind(self).(*args)
    end

    # Patches the original Axlsx::Package#to_stream method so that styles are
    # applied when the workbook is converted to StringIO
    original_to_stream = instance_method(:to_stream)
    define_method :to_stream do |*args|
      workbook.apply_styles if !workbook.styles_applied
      original_to_stream.bind(self).(*args)
    end
  end

  class Styles
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

    # Patches the original Axlsx::Styles#add_style method so that plain axlsx
    # styles are added to the axlsx_styler style_index cache
    original_add_style = instance_method(:add_style)
    define_method :add_style do |style|
      self.style_index ||= {}

      raw_style = {type: :xf, name: 'Arial', sz: 11, family: 1}.merge(style)
      if raw_style[:format_code]
        raw_style.delete(:num_fmt)
      end

      index = style_index.key(raw_style)
      if !index
        index = original_add_style.bind(self).(style)
        self.style_index[index] = raw_style
      end
      return index
    end

    private 

  end
end
