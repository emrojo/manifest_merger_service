class HeaderValidation
  def initialize(data_validation, header)
    @data_validation = data_validation
    @header = header
  end

  def formula_position
    header.cell.column_pos
  end

  def formula1
    "'#{@header.sheet.name}'!#{data_validation_range}"
  end

  def error
    "Only accepts values in (#{@data_validation.valid_values.join(',')})"
  end

  def valid_values
    @data_validation.valid_values
  end

  alias :prompt :error

  def data_validation_range
    @header.sheet.column_selection_to_str(index_to_column_name(@header.cell.column_pos), @sheet.row_pos, valid_values.length)
  end
end
