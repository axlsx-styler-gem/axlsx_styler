This gem is supposed to make styling spreadsheets easier.

Here's how the [example from the README](../README.md) compares to that implemented with plain `axlsx`.

```ruby
require 'axlsx'
axlsx = Axlsx::Package.new
wb = axlsx.workbook
border_color = '000000'
wb.add_worksheet do |sheet|
  # top row
  header_hash = { b: true, bg_color: '95AFBA' }
  top_left_corner = wb.styles.add_style header_hash.merge({
    border: { style: :thin, color: border_color, edges: [:top, :left, :bottom] }
  })
  top_edge = wb.styles.add_style header_hash.merge({
    border: { style: :thin, color: border_color, edges: [:top, :bottom] }
  })
  top_right_corner = wb.styles.add_style header_hash.merge({
    border: { style: :thin, color: border_color, edges: [:top, :right, :bottom] }
  })
  sheet.add_row
  sheet.add_row(["", "Product", "Category", "Price"],
    style: [ nil, top_left_corner, top_edge, top_right_corner ]
  )

  # middle rows
  color_hash = { bg_color: 'E2F89C' }
  left_edge = wb.styles.add_style color_hash.merge(
    b: true,
    border: {
      style: :thin, color: border_color, edges: [:left]
    }
  )
  inner = wb.styles.add_style color_hash
  right_edge = wb.styles.add_style color_hash.merge(
    alignment: { horizontal: :left },
    border: {
      style: :thin, color: border_color, edges: [:right]
    }
  )
  sheet.add_row(
    ["", "Butter", "Dairy", 4.99],
    style: [nil, left_edge, inner, right_edge]
  )
  sheet.add_row(
    ["", "Bread", "Baked Goods", 3.45],
    style: [nil, left_edge, inner, right_edge]
  )

  # last row
  bottom_left_corner = wb.styles.add_style color_hash.merge({
    b: true,
    border: { style: :thin, color: border_color, edges: [:left, :bottom] }
  })
  bottom_edge = wb.styles.add_style color_hash.merge({
    border: { style: :thin, color: border_color, edges: [:bottom] }
  })
  bottom_right_corner = wb.styles.add_style color_hash.merge({
    alignment: { horizontal: :left },
    border: { style: :thin, color: border_color, edges: [:right, :bottom] }
  })
  sheet.add_row(["", "Broccoli", "Produce", 2.99],
    style: [nil, bottom_left_corner, bottom_edge, bottom_right_corner]
  )

  sheet.column_widths 5, 20, 20, 20
end
axlsx.serialize "grocery.xlsx"
```
