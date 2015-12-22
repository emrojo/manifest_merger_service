require 'axlsx'
require_relative 'label_definitions'
require_relative 'excel_manifest/configuration_worksheet'
require_relative 'excel_manifest/data_worksheet'
require 'pry'

class ExcelBuilder


  def conditionals
    r = Axlsx::ConditionalFormattingRule.new(
    :formula => "=AND(ISBLANK(G10),COUNTA($C10:$BB10)>0)",
    :stopIfTrue => true)
    r.color_scale=Axlsx::ColorScale.new({:type => :num, :val => 0.55, :color => 'fff7696c'})
    [r]

    # 'Sample Registration'!$G$10:$G$9999
    #=AND(NOT(ISBLANK(G10)),NOT(OR(UPPER(G10)=$BK$3:$BK$8))) 'Sample Registration'!$G$10:$G$9999
  end

  def validation
    #allow list =$BK$3:$BK$8
  end

  def initialize(label_definitions)
    @label_definitions = label_definitions

  @p = Axlsx::Package.new
    @p.workbook.add_worksheet(:name => "-- NOT EDIT -- Configuration") do |sheet|
      @config_sheet = ExcelManifest::ConfigurationWorksheet.new(@label_definitions)
      @config_sheet.write_values_list(sheet)
      #@config_sheet.each_line do |line|
      #  @config_sheet.add_row line
      #end
      #config_sheet.add_row @label_definitions.conditional_formatting_lists
    end
    @p.workbook.add_worksheet(:name => "Basic Worksheet") do |sheet|
      @data_sheet = ExcelManifest::DataWorksheet.new(@label_definitions)
      @data_sheet.write_headers(sheet)
      @data_sheet.write_data_validations(sheet)


      #conditional_formattings(column_headers).each do |range,opts|
      #  sheet.add_conditional_formatting(range, opts)
      #end
      #sheet.add_conditional_formatting("A2:A100", {
      #  :formula => "=AND(ISBLANK(G10),COUNTA($C10:$BB10)>0)",
      # :stopIfTrue => true})
      #sheet.add_data_validation("A2:A10", {
      # :type => :list,
      #  :formula1 => 'C1:C8',
      #  :showDropDown => false,
      #  :showErrorMessage => true,
      #  :errorTitle => '',
      #  :error => 'Only values from C1:C8',
      #  :errorStyle => :stop,
      #  :showInputMessage => true,
      #  :promptTitle => '',
      #  :prompt => 'Only values from C1:C8'
      #})
    end
    @p.use_shared_strings = true
  end
  def serialize
    @p.serialize('simple.xlsx')
  end
end
ExcelBuilder.new(LabelDefinitions.new).serialize
