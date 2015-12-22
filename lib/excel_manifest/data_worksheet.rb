class ExcelManifest::DataWorksheet

  include ExcelManifest

  @@COLUMN_POS = 1
  @@ROW_VALUES_LIST_START = 1

  def field_column_name(field)
    index_to_column_name(@label_definitions.column_headers.find_index(field))
  end

  def field_column_range(field)
    column_selection_to_str(field_column_name(field), DEFAULT_ROW_RANGE.first, DEFAULT_ROW_RANGE.count)
  end

  def write_headers(sheet)
    sheet.add_row @label_definitions.column_headers
  end

  def write_data_validations(sheet)
    @label_definitions.column_headers.each do |field|
      validation_options = @label_definitions.data_validation_options(field)
      sheet.add_data_validation(field_column_range(field), validation_options) unless validation_options.nil?
    end
  end

  def initialize(label_definitions)
    @label_definitions = label_definitions
  end
end
