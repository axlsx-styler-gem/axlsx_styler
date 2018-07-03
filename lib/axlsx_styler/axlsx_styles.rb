if RUBY_VERSION.to_f >= 2

  module AxlsxStyler
    module Styles

      # An index for cell styles
      #   {
      #     1 => < style_hash >,
      #     2 => < style_hash >,
      #     ...
      #     K => < style_hash >
      #   }
      # where keys are Cell#raw_style and values are styles codes as per Axlsx::Style
      attr_accessor :style_index

      # Ensure plain axlsx styles are added to the axlsx_styler style_index cache
      def add_style(*args)
        style = args.first

        self.style_index ||= {}

        raw_style = {type: :xf, name: 'Arial', sz: 11, family: 1}.merge(style)
        if raw_style[:format_code]
          raw_style.delete(:num_fmt)
        end

        index = style_index.key(raw_style)
        if !index
          index = super 
          self.style_index[index] = raw_style
        end
        return index
      end

    end
  end

  Axlsx::Styles.send(:prepend, AxlsxStyler::Styles)

else

  module Axlsx
    class Styles

      # An index for cell styles
      #   {
      #     1 => < style_hash >,
      #     2 => < style_hash >,
      #     ...
      #     K => < style_hash >
      #   }
      # where keys are Cell#raw_style and values are styles codes as per Axlsx::Style
      attr_accessor :style_index

      # Ensure plain axlsx styles are added to the axlsx_styler style_index cache
      original_add_style = instance_method(:add_style)
      define_method :add_style do |*args|
        style = args.first

        self.style_index ||= {}

        raw_style = {type: :xf, name: 'Arial', sz: 11, family: 1}.merge(style)
        if raw_style[:format_code]
          raw_style.delete(:num_fmt)
        end

        index = style_index.key(raw_style)
        if !index
          index = original_add_style.bind(self).(*args)
          self.style_index[index] = raw_style
        end
        return index
      end

    end
  end

end
