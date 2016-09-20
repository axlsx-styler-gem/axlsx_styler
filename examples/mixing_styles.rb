# As of v0.1.7 it is possible to add styles using vanilla axlsx approach
# as well as using add_style and add_border. This is something users have
# been asking for.
#
# The example below is superfluous and the same results are more easily achieved
# without workbook.styles.add_style. You've been warned!
require 'axlsx_styler'

axlsx = Axlsx::Package.new
workbook = axlsx.workbook
red = workbook.styles.add_style fg_color: 'FF0000'
workbook.add_worksheet do |sheet|
  sheet.add_row
  sheet.add_row ['', 'B2', 'C2', 'D2'], style: [nil, red, nil, nil]
  sheet.add_row ['', 'B3', 'C3', 'D3'], style: [nil, red, nil, nil]
  sheet.add_row ['', 'B4', 'C4', 'D4'], style: [nil, red, nil, nil]

  sheet.add_style 'B2:D2', b: true
  sheet.add_border 'B2:D4'
end
workbook.apply_styles
axlsx.serialize File.expand_path('../../tmp/mixing_styles.xlsx', __FILE__)
