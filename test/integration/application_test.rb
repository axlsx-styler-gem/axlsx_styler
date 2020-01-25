### FOR RAILS TESTS

require 'test_helper'

class ApplicationTest < ActionDispatch::IntegrationTest

  def test_xlsx
    get '/spreadsheets/test.xlsx'
    assert_response :success

    path = File.expand_path("../../../tmp/axlsx_rails.xlsx", __FILE__)

    File.open(path, 'w+b') do |f|
      f.write @response.body
    end
  end

end
