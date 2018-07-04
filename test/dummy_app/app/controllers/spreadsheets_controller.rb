class SpreadsheetsController < ApplicationController

  def xlsx
    render xlsx: "spreadsheets/xlsx", layout: false
  end

end
