require 'sinatra'
require File.expand_path(File.dirname(__FILE__) +'/../lib/excel_merge')

post '/csv_to_spreadsheet' do
  ExcelBuilder.new(params['csv']).serialize
end

post '/merge_excel/' do
  xls_template = params['template'][:tempfile]
  samples_csv  = params['manifest-details'][:tempfile]
  manifest     = ExcelMerge.merge_template_with_csv(xls_template, samples_csv)

  manifest.write(manifest.io)

  # Use these headers when running as localhost.  Not needed behind nginx though...
  # response.headers['content_type'] = "application/vnd.ms-excel"
  # response.headers['Access-Control-Allow-Origin'] = '*'
  # response.headers['Access-Control-Allow-Methods'] = 'OPTIONS,GET,HEAD,POST,PUT,DELETE,TRACE,CONNECT'

  send_file(manifest.io)
end

