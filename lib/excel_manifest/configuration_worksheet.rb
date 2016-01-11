

class ConfigurationWorksheet < Worksheet
  DEFAULT_COLUMN_POS = 4
  DEFAULT_ROW_POS = 5


  attr_accessor :column_pos
  attr_accessor :row_pos

  def initialize(package, name, headers)
    @column_pos = DEFAULT_COLUMN_POS
    @row_pos = DEFAULT_ROW_POS
    @INDEX = @column_pos

    super(package, name, headers)

    @package.workbook.add_worksheet(:name => @name, :state => :hidden) do |sheet|
      @sheet = sheet
      write_data_validation_configuration
      lock_worksheet_with_password('fish')
    end
  end

  def data_validation_formula(header)
    "'#{@sheet_name}'!#{data_validation_list_column_range(formula_position, valid_values)}"
  end

  def data_validations
    @headers.map(&:data_validation)
      dv.formula_position = idx + @column_pos
      dv.formula1 = data_validation_formula(header)
      DataValidation.find(:all).select{|dv| (dv.labels * @labels).length > 0 }.map do |dv|

    end
  end

  def write_data_validation_configuration
    update_sheet_with_table(@sheet, transpose_table(@headers.map(&:data_validation).map(&:valid_values)), @row_pos, @column_pos)

    # idx_column_pos = @column_pos-1
    # @label_definitions.definitions_with_data_validation.map do |label_definition|
    #   update_label_definition(label_definition, idx_column_pos)
    #   idx_column_pos += 1
    #   label_definition[:data_validation][:valid_values]
    # end.tap do |table|
    #   update_sheet_with_table(@sheet, transpose_table(table), @row_pos, @column_pos)
    # end
  end
end

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
