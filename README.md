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

## Installation

Add this line to your application's Gemfile:

    gem 'axlsx_styler'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install axlsx_styler
