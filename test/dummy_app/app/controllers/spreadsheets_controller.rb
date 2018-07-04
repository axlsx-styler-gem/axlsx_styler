class SpreadsheetsController < ApplicationController

  def xlsx
    respond_to do |format|
      format.xlsx
    end
  end

end
