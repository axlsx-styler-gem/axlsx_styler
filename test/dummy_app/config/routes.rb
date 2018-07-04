Dummy::Application.routes.draw do
  get 'spreadsheet/xlsx', to: 'spreadsheets#xlsx'
end
