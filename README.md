[![Build Status](https://travis-ci.org/sakovias/axlsx_styler.svg?branch=master)](https://travis-ci.org/sakovias/axlsx_styler)

# axlsx_styler

[axlsx](https://github.com/randym/axlsx) gem is an excellent tool to
build Excel spreadsheets. The sheets are
created row-by-row and styles are immediately added to each cell when a
row is created.

`axlsx_styler` allows to separate styles from content: you can fill out
a spreadsheet with data and apply styles later. Paired with
[axlsx_rails](https://github.com/straydogstudio/axlsx_rails) this gem
allows to build clean and maintainable Excel views in a Rails app. It can also
be used outside of any specific Ruby framework as shown in example below.

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
sheet.add_border 'B2:D5', [:bottom, :right]
sheet.add_border 'B2:D5', { edges: [:bottom, :right], style: :thick, color: 'FF0000' }
```

Border parameters are optional. The default is to draw a thin black border on all four edges of the selected cell range.


### Example

Suppose we want create the following spreadsheet:

![alt text](./spreadsheet.png "Sample Spreadsheet")

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
  sheet.add_border 'B3:D3', [:top]
end
axlsx.serialize 'grocery.xlsx'
```

Producing the same spreadsheet with vanilla `axlsx` turns out [a bit trickier](./examples/vanilla_axlsx.md).


## Change log

Version | Change
--------|-------
0.1.7 | Allow mixing vanilla `axlsx` styles and those from `add_style` and `add_border` (see [this example](./examples/mixing_styles.rb))
0.1.5 | Hide `Workbook#apply_styles` method to make it easier to use.
0.1.3 | Make border styles customazible.
0.1.2 | Allow applying multiple style hashes.
