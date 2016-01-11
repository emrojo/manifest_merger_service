class DataWorksheet < Worksheet
  DEFAULT_ROW_POS = 2
  DEFAULT_COLUMN_POS = 3
  DATA_ROW_SIZE = 99

  def initialize(package, name, headers)
    @column_pos = DEFAULT_COLUMN_POS
    @row_pos = DEFAULT_ROW_POS

    @INDEX = @column_pos

    super(package, name, headers)

    package.workbook.add_worksheet(:name => "Sample Registration") do |sheet|
      @sheet = sheet
      write_headers
      write_conditional_formatting(@config_sheet)
      #write_data_validations
    end
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

  def write_headers
    update_sheet_with_table(@sheet,[@headers.map(&:name).map(&:name)], @row_pos, @column_pos)
    #fill_sheet_with_empty_rows_until_row_pos(@sheet, @row_pos)
    #@sheet.add_row @label_definitions.column_headers
  end

  def write_conditional_formatting(config_sheet)
    styles_created = Style.all.map{|s| WorkbookStyle.new(@sheet.workbook, s)}

    @headers.map(&:conditional_formattings).each do |cf_list|
      cf_list.each do |cf|
        Axlsx::ConditionalFormattingRule.new(:type => :expression,
                                                      :priority => 1,
                                                      :dxfId => styles_created.select{|s| s[:name] == cf[:style]}.first.style_id,
                                                      :formula => cf.formula)

        @sheet.add_conditional_formatting(field_column_range(label_name), [cfr])
      end
    end
    # @label_definitions.conditional_formattings.each do |cf|
    #   cf[:labels].each do |label_name|
    #     formula = @label_definitions.rules.select{|r| r[:name] == cf[:rule]}.first[:formula]
    #     created_formula = CellContext.new(self, config_sheet, label_name).create_formula(formula)
    #     cfr = Axlsx::ConditionalFormattingRule.new( { :type => :expression,
    #                                                   :priority => 1,
    #                                                   :dxfId => styles_created.select{|s| s[:name] == cf[:style]}.first.style_id,
    #                                                   :formula => created_formula })
    #     @sheet.add_conditional_formatting(field_column_range(label_name), [cfr])
    #   end
    # end

  end

end


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
    column_selection_to_str(field_column_name(field), data_row_range.first, data_row_range.count)
  end

  def field_column_range_selection(field)
    column_selection_to_str_selection(field_column_name(field), data_row_range.first, data_row_range.count)
  end


  def write_headers
    update_sheet_with_table(@sheet, [@label_definitions.column_headers], @row_pos, @column_pos)
    #fill_sheet_with_empty_rows_until_row_pos(@sheet, @row_pos)
    #@sheet.add_row @label_definitions.column_headers
  end


  class CellContext
    def initialize(sheet, config_sheet, name)
      @sheet = sheet
      @config_sheet = config_sheet
      @name = name
    end

    def get_binding
      return binding()
    end

    def create_formula(formula)
      eval("%Q{#{formula}}", get_binding)
    end

    def my_cell
      @sheet.field_column_range(@name)
      @sheet.label_definition(@name)
      "#{@sheet.field_column_name(@name)}#{@sheet.data_row_range.first+1}"
    end

    def range_of_entry
      @sheet.field_column_range_selection(@name)
    end

    def range_of_validation
      @config_sheet.data_validation_formula(@sheet.label_definition(@name))
    end
  end

  def label_definition(name)
    @label_definitions.label_definition(name)
  end

  def write_conditional_formatting(config_sheet)

    styles_created = @label_definitions.styles.map do |style|
      { :name => style[:name],
        :st_created => @sheet.workbook.styles.add_style(style.tap{|s| s.delete(:name)})
      }
    end
    @label_definitions.conditional_formattings.each do |cf|
      cf[:labels].each do |label_name|
        formula = @label_definitions.rules.select{|r| r[:name] == cf[:rule]}.first[:formula]
        created_formula = CellContext.new(self, config_sheet, label_name).create_formula(formula)
        cfr = Axlsx::ConditionalFormattingRule.new( { :type => :expression,
                                                      :priority => 1,
                                                      :dxfId => styles_created.select{|s| s[:name] == cf[:style]}.first[:st_created],
                                                      :formula => created_formula })
        @sheet.add_conditional_formatting(field_column_range(label_name), [cfr])
      end
    end

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
