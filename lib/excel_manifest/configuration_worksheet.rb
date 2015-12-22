class ExcelManifest::ConfigurationWorksheet

  include ExcelManifest

  @@COLUMN_POS = 1
  @@ROW_VALUES_LIST_START = 1

  def validation_list_column_range(pos, list)
    column_selection_to_str(index_to_column_name(pos), @@ROW_VALUES_LIST_START, list.length)
  end

  def formula(label_definition)
    "'-- NOT EDIT -- Configuration'!#{validation_list_column_range(label_definition[:data_validation][:formula_position], label_definition[:data_validation][:valid_values])}"
  end

  def update_validation_lists!
    @label_definitions.each_label_definition do |label_definition|
      if label_definition[:type]==:list
        @@COLUMN_POS += 1
        label_definition[:data_validation][:formula_position] = @@COLUMN_POS
        label_definition[:data_validation][:formula1] = formula(label_definition)
      end
    end
  end

  def write_values_list(sheet)

  end

  def initialize(label_definitions)
    @label_definitions = label_definitions
    update_validation_lists!
  end

end
