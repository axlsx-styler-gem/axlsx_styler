# Changelog

- **1.0.0 - UNRELEASED**
  - Improve Package and Styles monkey patches using `prepend` for Ruby 2+, reverts to old patch when using Ruby `1.9.3`.
  - Removed unnecessary module `AxlsxStyler::Axlsx`.
  - Major test suite improvements. Add Rails to test suite. Add axlsx_master to appraisals. Test all minor Ruby versions down to `v1.9.3`.
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
