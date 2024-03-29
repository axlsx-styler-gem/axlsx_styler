# Changelog

- **Unreleased**
  - Nothing yet

- **1.2.0 - October 21, 2022**
  - This gem has now been merged into caxlsx v3.3.0 as such we are now deprecating this gem and dropping support for caxlsx v3.3.0+
  - [#35](https://github.com/axlsx-styler-gem/axlsx_styler/pull/35) - Do not overwrite spreadsheet default font attributes if they have been customized

- **1.1.0 - August 26, 2020**
  - [Issue #29](https://github.com/axlsx-styler-gem/axlsx_styler/issues/29) - Fix error `Invalid cellXfs id` when applying `dxf` styles
  - [PR #28](https://github.com/axlsx-styler-gem/axlsx_styler/pull/28) - Allow passing arrays of cell ranges to the `add_style` and `add_border` methods

- **1.0.0 - January 5, 2020**
  - Switch to the `caxlsx` gem (Community Axlsx) from the legacy unmaintained `axlsx` gem. Axlsx has had a long history of being poorly maintained so this community gem improves the situation.
  - Require Ruby 2.3+
  - Improve Package and Styles monkey patches using `prepend`
  - Removed unnecessary module `AxlsxStyler::Axlsx`
  - Major test suite improvements
  - Add Rails dummy app to test suite and test with `axlsx_rails`

- **0.2.0**
  - Add support for `axlsx` v3 (at this time for v3.0.0.pre).
  - Update test suite to run against both v2 and v3 of `axlsx`

- **0.1.7**
  - Allow mixing vanilla `axlsx` styles and those from `add_style` and `add_border`. See [this example](./examples/mixing_styles.rb)

- **0.1.5**
  - Hide `Workbook#apply_styles` method to make it easier to use.

- **0.1.3**
  - Make border styles customizable.

- **0.1.2**
  - Allow applying multiple style hashes.
