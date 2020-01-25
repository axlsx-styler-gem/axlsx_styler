# Changelog

- **1.0.0 - UNRELEASED**
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
