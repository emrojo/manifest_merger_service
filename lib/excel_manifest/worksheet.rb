class Worksheet
  include ExcelManifest
  attr_accessor :package, :name, :headers

  def initialize(package, name, headers)
    @package = package
    @name = name
    @INDEX = 0
    @headers = headers.map do |header_name|
      Header.new(header_name, @name, get_new_header_cell)
    end
  end

  def get_new_header_cell
    @INDEX = @INDEX + 1
    Cell.new(@row_pos, @column_pos + @INDEX)
  end

  def lock_worksheet_with_password(password)
    @sheet.sheet_protection.password = 'fish'
  end


end
