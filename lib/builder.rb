require 'active_record'
require 'pry'

require_relative 'label_definitions'
require_relative 'excel_manifest/rule'
require_relative 'excel_manifest/style'
require_relative 'excel_manifest/manifest_template'
require_relative 'excel_manifest/header'
require_relative 'excel_manifest/worksheet'
require_relative 'excel_manifest/configuration_worksheet'
require_relative 'excel_manifest/data_worksheet'
require_relative 'excel_manifest/data_validation'
require_relative 'excel_manifest/header_validation'
require_relative 'excel_manifest/style'
require_relative 'excel_manifest/workbook_style'
require_relative 'excel_manifest/cell'
require_relative 'excel_manifest/conditional_formatting'


ConditionalFormatting.all = LabelDefinitions.new.conditional_formattings.map{|params| ConditionalFormatting.new(params)}
LabelDefinitions.new.rules.map{|params| Rule.new(params)}
Style.all = LabelDefinitions.new.styles.map{|params| Style.new(params)}
DataValidation.all = LabelDefinitions.new.data_validations.map{|params| DataValidation.new(params)}
ManifestTemplate.all = LabelDefinitions.new.manifest_templates.map{|params| ManifestTemplate.new(params)}




ManifestTemplate.all.first.serialize
