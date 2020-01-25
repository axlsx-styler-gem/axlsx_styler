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
