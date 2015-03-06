require 'axlsx'

require 'axlsx_styler/version'
require 'axlsx_styler/axlsx_workbook'
require 'axlsx_styler/axlsx_worksheet'
require 'axlsx_styler/axlsx_cell'

Axlsx::Workbook.send :include, AxlsxStyler::Axlsx::Workbook
Axlsx::Worksheet.send :include, AxlsxStyler::Axlsx::Worksheet
Axlsx::Cell.send :include, AxlsxStyler::Axlsx::Cell
