class Header
  attr_accessor :name, :sheet, :cell

  def initialize(name, sheet, cell)
    @name = name
    @sheet = sheet
    @cell = cell
  end

  def data_validation
    DataValidation.all.select{|dv| dv.headers.include?(name)}.first.header_validation(self)
  end

  def sheet_name
    @sheet.name
  end

  def conditional_formattings
    ConditionalFormatting.all.select{|cf| cf.labels.include?(@name)}
  end

end
