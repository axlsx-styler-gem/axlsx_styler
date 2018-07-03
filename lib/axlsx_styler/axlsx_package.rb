if RUBY_VERSION.to_f >= 2

  module AxlsxStyler
    module Package

      def serialize(*args)
        workbook.apply_styles if !workbook.styles_applied
        super
      end

      def to_stream(*args)
        workbook.apply_styles if !workbook.styles_applied
        super
      end

    end
  end

  Axlsx::Package.send(:prepend, AxlsxStyler::Package)

else

  module Axlsx
    class Package
      #  Apply styles when the workbook is saved
      original_serialize = instance_method(:serialize)
      define_method :serialize do |*args|
        workbook.apply_styles if !workbook.styles_applied
        original_serialize.bind(self).(*args)
      end

      # Apply styles when the workbook is converted to StringIO
      original_to_stream = instance_method(:to_stream)
      define_method :to_stream do |*args|
        workbook.apply_styles if !workbook.styles_applied
        original_to_stream.bind(self).(*args)
      end
    end
  end

end
