require 'test_helper'

class ExamplesTest < MiniTest::Test

  def setup
    @axlsx = Axlsx::Package.new
    @workbook = @axlsx.workbook
  end
  
  def teardown
  end

  def test_examples
    path = File.expand_path("../../../examples/**.rb", __FILE__)
    Dir[path].each do |file| 
      require file
    end
  end

end
