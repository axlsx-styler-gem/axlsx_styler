module AxlsxStyler
  module Styles

    # An index for cell styles where keys are styles codes as per Axlsx::Style and values are Cell#raw_style
    # The reason for the backward key/value ordering is that style lookup must be most efficient, while `add_style` can be less efficient
    def style_index
      @style_index ||= {}
    end

    # Ensure plain axlsx styles are added to the axlsx_styler style_index cache
    def add_style(options={})
      if options[:type] == :dxf
        style_id = super
      else
        ### https://github.com/caxlsx/caxlsx/blob/9b6a78f43b9415bd155bddf6a6d872f5d5555595/lib/axlsx/stylesheet/styles.rb#L459
        font_defaults = {name: @fonts.first.name, sz: @fonts.first.sz, family: @fonts.first.family} 

        raw_style = {type: :xf}.merge(font_defaults).merge(options)

        if raw_style[:format_code]
          raw_style.delete(:num_fmt)
        end

        style_id = style_index.key(raw_style)

        if !style_id
          style_id = super 

          style_index[style_id] = raw_style
        end
      end

      return style_id
    end

  end
end

Axlsx::Styles.send(:prepend, AxlsxStyler::Styles)
