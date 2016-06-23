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
      @workbook.apply_styles if !@workbook.styles_applied
      original_serialize.bind(self).(*args)
    end

    # Patches the original Axlsx::Package#to_stream method so that styles are
    # applied when the workbook is converted to StringIO
    original_to_stream = instance_method(:to_stream)
    define_method :to_stream do |*args|
      @workbook.apply_styles if !@workbook.styles_applied
      original_to_stream.bind(self).(*args)
    end
  end
end
