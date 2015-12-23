class ExcelManifest::DataWorksheet

  DEFAULT_ROW_POS = 2
  DEFAULT_COLUMN_POS = 3
  DATA_ROW_SIZE = 99

  include ExcelManifest

  attr_accessor :column_pos

  def field_column_name(field)
    index_to_column_name(@label_definitions.column_headers.find_index(field)+@column_pos-1)
  end

  def field_column_range(field)
    column_selection_to_str(field_column_name(field), DATA_ROW_RANGE.first, DATA_ROW_RANGE.count)
  end

  def write_headers
    update_sheet_with_table(@sheet, [@label_definitions.column_headers], @row_pos, @column_pos)
    #fill_sheet_with_empty_rows_until_row_pos(@sheet, @row_pos)
    #@sheet.add_row @label_definitions.column_headers
  end

  def write_conditional_formatting

  end

  def data_row_range
    (@row_pos..(@row_pos + DATA_ROW_SIZE))
  end

  def write_data_validations
    @label_definitions.column_headers.each do |field|
      validation_options = @label_definitions.data_validation_options(field)
      data_row_range.each do |pos|
        @sheet.add_data_validation("#{field_column_name(field)}#{pos+@row_pos}", validation_options) unless validation_options.nil?
      end
      #sheet.add_data_validation(field_column_range(field), validation_options) unless validation_options.nil?
    end
  end

  def reset
    @column_pos = DEFAULT_COLUMN_POS
    @row_pos = DEFAULT_ROW_POS
  end

  def initialize(sheet, label_definitions)
    @column_pos = 1
    @label_definitions = label_definitions
    @sheet = sheet
    reset
  end
end
