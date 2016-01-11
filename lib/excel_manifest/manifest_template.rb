require 'axlsx'

class ManifestTemplate
  DEFAULT_CONFIGURATION_WORKSHEET_NAME = "-- NOT EDIT -- Configuration"
  DEFAULT_DATA_WORKSHEET_NAME = "Sample Registration"

  attr_accessor :headers
  attr_accessor :data_worksheet
  attr_accessor :configuration_worksheet

  @@all = nil

  def self.all
    @@all
  end

  def self.all=(value)
    @@all=value
  end

  def initialize(params)
    @key = params[:key]
    @headers = params[:headers]

    @p = Axlsx::Package.new
    @configuration_worksheet = ConfigurationWorksheet.new(@p, DEFAULT_CONFIGURATION_WORKSHEET_NAME, @headers)
    #binding.pry
    @data_worksheet = DataWorksheet.new(@p, DEFAULT_DATA_WORKSHEET_NAME, @configuration_worksheet.headers)
    @p.use_shared_strings = true
  end

  def serialize
    @p.serialize('simple.xlsx')
  end

end
