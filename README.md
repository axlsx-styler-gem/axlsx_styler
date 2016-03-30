# axlsx_styler

[axlsx](https://github.com/randym/axlsx) gem is an excellent tool to
build Excel spreadsheets. The sheets are
created row-by-row and styles are immediately added to each cell when a
row is created.

`axlsx_styler` allows to separate styles from content: you can fill out
a spreadsheet with data and apply styles later. Paired with
[axlsx_rails](https://github.com/straydogstudio/axlsx_rails) this gem
allows to build clean and maintainable Excel views in a Rails app. It can also
be used outside of any specific ruby framework as shown in example below.

## Usage

This gem provides a DSL that allows you to apply styles to ranges of cells, e.g.

```ruby
sheet.add_style 'A1:D10', b: true, sz: 14
```

The styles can be overlayed, so that later on you can add another style
to cells that already have styles, e.g.

```ruby
sheet.add_style 'A1:D1', bg_color: 'FF0000'
```

Applying multiple styles as a sequence of Ruby hashes is also possible:

```ruby
bold     = { b: true }
centered = { alignment: { horizontal: :center } }
sheet.add_style 'A2:D2', bold, centered
```

You can also add borders to ranges of cells:

```ruby
sheet.add_border 'B2:D5'
sheet.add_border 'B2:D5', [:right]
sheet.add_border 'B2:D5', [:right], :thick
border_color = 'FF0000'
sheet.add_border 'B2:D5', [:right], :medium, border_color
```

When you are done with styling you just need to run

```ruby
workbook.apply_styles
```

Here's an example that compares styling a simple table with and without
`axlsx_styler`. Suppose we wand to create the following spreadsheet:

![alt text](./spreadsheet.png "Sample Spreadsheet")

### `axlsx` paired with `axlsx_styler`

You can apply styles after all data is entered, similar to how you'd create
an Excel document by hand:

```ruby
require 'axlsx_styler'

axlsx = Axlsx::Package.new
workbook = axlsx.workbook
workbook.add_worksheet do |sheet|
  sheet.add_row
  sheet.add_row ['', 'Product', 'Category',  'Price']
  sheet.add_row ['', 'Butter', 'Dairy',      4.99]
  sheet.add_row ['', 'Bread', 'Baked Goods', 3.45]
  sheet.add_row ['', 'Broccoli', 'Produce',  2.99]
  sheet.column_widths 5, 20, 20, 20

  # using AxlsxStyler DSL
  sheet.add_style 'B2:D2', b: true
  sheet.add_style 'B2:B5', b: true
  sheet.add_style 'B2:D2', bg_color: '95AFBA'
  sheet.add_style 'B3:D5', bg_color: 'E2F89C'
  sheet.add_style 'D3:D5', alignment: { horizontal: :left }
  sheet.add_border 'B2:D5'
  sheet.add_border 'B3:D3', [:top], :medium
end
workbook.apply_styles
axlsx.serialize 'grocery.xlsx'
```

### `axlsx` gem without `axlsx_styler`

Whith plain `axlsx` you need to know which styles you're going to use beforehand.
The code for our example is a bit more envolved:

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
