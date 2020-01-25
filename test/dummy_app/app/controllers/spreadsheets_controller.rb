class SpreadsheetsController < ApplicationController

  def test
    render xlsx: "test", layout: false
  end

end
