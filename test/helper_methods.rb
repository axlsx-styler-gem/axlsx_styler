# Save to a file using Axlsx::Package#serialize
def serialize(filename)
  @axlsx.serialize File.expand_path("../../tmp/#{filename}.xlsx", __FILE__)
  assert_equal true, @workbook.styles_applied
end

# Save to a file by getting contents from stream
def to_stream(filename)
  File.open(File.expand_path("../../tmp/#{filename}.xlsx", __FILE__), 'w') do |f|
    f.write @axlsx.to_stream.read
  end
end
