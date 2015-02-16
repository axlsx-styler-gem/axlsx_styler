# AxlsxStyler

[Axlsx](https://github.com/randym/axlsx) gem is an excellent tool to
build Excel spreadsheets. The sheets are
created row-by-row and styles are immediately added to each cell when a
row is created. This gem allows to follow an alternative route: fill out
a spreadsheet with data and apply styles later. Styles can be added
to individual cells as well as to ranges of cells. As a bonus, this gem
also simplifies drawing borders around groups of cells.

## Usage

This gem extend `Array` class in a way that allows you to apply
styles to ranges of cells, e.g.

```ruby
sheet["A1:D10"].add_style b: true
```

The styles can be overlayed, so that later on you can add another style
to cells that already have styles, e.g.

```ruby
sheet["A1:D1"].add_style bg_color: 'FF0000'
```

You can also add borders to ranges of cells:

```ruby
sheet["B2:D5"].add_border
sheet["B2:B5"].add_border [:right]
```

When you are done with styling you just need to run

```ruby
workbook.apply_styles
```

Here's an example that compares styling a simple table with and without
AxlsxStyler. Suppose make our spreadsheet to look as follows:

![alt text](./spreadsheet.png "Sample Spreadsheet")

### With `AxlsxStyler`

Just follow the step outlined above:

```ruby
require 'axlsx_styler'

axlsx = Axlsx::Package.new
workbook = axlsx.workbook

workbook.add_worksheet do |sheet|
  sheet.add_row
  sheet.add_row ["", "Product", "Category", "Price"]
  sheet.add_row ["", "Butter", "Dairy", 4.99]
  sheet.add_row ["", "Bread", "Baked Goods", 3.45]
  sheet.add_row ["", "Broccoli", "Produce", 2.99]

  sheet["B2:D2"].add_style b: true
  sheet["B2:B5"].add_style b: true
  sheet["B2:D2"].add_style bg_color: "95AFBA"
  sheet["B3:D5"].add_style bg_color: "E2F89C"
  sheet["D3:D5"].add_style alignment: { horizontal: :left }
  sheet["B2:D5"].add_border
  sheet["B3:D3"].add_border [:top]
  sheet.column_widths 5, 20, 20, 20
end

workbook.apply_styles

axlsx.serialize "grocery.xlsx"
```

### With plain `Axlsx` gem

This example can be DRYied up, but it gives you a rough idea of what you
need to go through.

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

## Installation

Add this line to your application's Gemfile:

    gem 'axlsx_styler'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install axlsx_styler
