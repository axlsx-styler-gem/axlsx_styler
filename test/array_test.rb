require 'test_helper'

class ArrayTest < MiniTest::Test
  def test_presesence_of_add_style
    array = Array.new
    assert array.respond_to?(:add_style)
  end
end
