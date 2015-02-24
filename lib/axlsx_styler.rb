require 'axlsx'

require 'axlsx_styler/version'
require 'axlsx_styler/array'
require 'axlsx_styler/axlsx_workbook'
require 'axlsx_styler/axlsx_cell'

Array.send :include, AxlsxStyler::Array
Axlsx::Workbook.send :include, AxlsxStyler::Axlsx::Workbook
Axlsx::Cell.send :include, AxlsxStyler::Axlsx::Cell
