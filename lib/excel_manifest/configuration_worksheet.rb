class ExcelManifest::ConfigurationWorksheet

  include ExcelManifest

  DEFAULT_COLUMN_POS = 4
  DEFAULT_ROW_POS = 5

  attr_accessor :column_pos
  attr_accessor :row_pos

  def data_validation_list_column_range(pos, list)
    column_selection_to_str(index_to_column_name(pos), @row_pos, list.length)
  end

  def data_validation_formula(label_definition)
    "'#{@sheet.name}'!#{data_validation_list_column_range(label_definition[:data_validation][:formula_position], label_definition[:data_validation][:valid_values])}"
  end

  def data_validation_error_msg(label_definition)
    "Only accepts values in (#{label_definition[:data_validation][:valid_values].join(',')})"
  end

  alias :data_validation_prompt_msg :data_validation_error_msg

  def update_label_definition(label_definition, idx_column_pos)
    label_definition[:data_validation][:formula_position] = idx_column_pos
    label_definition[:data_validation][:formula1] = data_validation_formula(label_definition)
    label_definition[:data_validation][:error] = data_validation_error_msg(label_definition)
    label_definition[:data_validation][:prompt] = data_validation_prompt_msg(label_definition)
  end

  def write_data_validation_configuration
    idx_column_pos = @column_pos
    @label_definitions.definitions_with_data_validation.map do |label_definition|
      update_label_definition(label_definition, idx_column_pos)
      idx_column_pos += 1
      label_definition[:data_validation][:valid_values]
    end.tap do |table|
      update_sheet_with_table(@sheet, transpose_table(table), @row_pos, @column_pos)
    end
  end

  def lock_worksheet_with_password(password)
    @sheet.sheet_protection.password = 'fish'
  end

  def reset
    @column_pos = DEFAULT_COLUMN_POS
    @row_pos = DEFAULT_ROW_POS
  end

  def initialize(sheet, label_definitions)
    @sheet = sheet
    @label_definitions = label_definitions
    reset
  end

end
