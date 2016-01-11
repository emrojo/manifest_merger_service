class DataValidation
  attr_accessor :headers, :valid_values

  @@all = nil

  def self.all
    @@all
  end

  def self.all=(value)
    @@all=value
  end


  def list_column_range(pos, list)
    @sheet.column_selection_to_str(index_to_column_name(pos), @row_pos, list.length)
  end

  def header_validation(header)
    HeaderValidation.new(self, header)
  end

  def initialize(params)
    @headers = params[:labels]
    @valid_values = params[:valid_values]
  end

end
