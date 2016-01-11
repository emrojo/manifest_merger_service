require_relative 'excel_manifest'

class LabelDefinitions

  include ExcelManifest

  def rules
    [
      {
        :key => :incorrect_format,
        :formula => %q{OR(LEN(#{my_cell})>15,NOT(ISERROR(FIND("&amp;",#{my_cell}))),NOT(ISERROR(FIND("(",#{my_cell}))),NOT(ISERROR(FIND(")",#{my_cell}))),NOT(ISERROR(FIND(" ",#{my_cell}))),NOT(ISERROR(FIND("+",#{my_cell}))),NOT(ISERROR(FIND("*",#{my_cell}))),NOT(ISERROR(FIND("-",#{my_cell}))))}
      },
      {
        :key => :incorrect_content,
        :formula => %q{OR(NOT(ISERROR(FIND(".",#{my_cell}))),NOT(ISERROR(FIND(",",#{my_cell}))),NOT(ISERROR(FIND("\\",#{my_cell}))),NOT(ISERROR(FIND("\/",#{my_cell}))))}
      },
      {
        :key => :empty_entry,
        :formula => %q{AND(ISBLANK(#{my_cell}),COUNTA(#{range_of_entry})&gt;0)}
      },
      {
        :key => :date_validation,
        :formula => %q{OR(AND(NOT(ISBLANK(#{my_cell})),NOT(LEN(#{my_cell})=5),NOT(ISERROR(FIND("/",#{my_cell})))),AND(NOT(ISBLANK(#{my_cell})),NOT(LEN(#{my_cell})=4),(ISERROR(FIND("/",#{my_cell})))),NOT(ISERROR(FIND(" ",#{my_cell}))))}
      },
      {
        :key => :number_validation,
        :formula => %q{AND(NOT(ISNUMBER(#{my_cell})),NOT(ISBLANK(#{my_cell})))}
      },
      {
        :key => :string_validation_30,
        :formula => %q{OR(LEN(#{my_cell})&gt;30,NOT(ISERROR(FIND("&amp;",#{my_cell}))),NOT(ISERROR(FIND("(",#{my_cell}))),NOT(ISERROR(FIND(")",#{my_cell}))),NOT(ISERROR(FIND("$",#{my_cell}))),NOT(ISERROR(FIND("*",#{my_cell}))),NOT(ISERROR(FIND("-",#{my_cell}))))"}
      },
      {
        :key => :string_validation_7,
        :formula => %q{OR(LEN(#{my_cell})&gt;7,NOT(ISERROR(FIND("&amp;",#{my_cell}))),NOT(ISERROR(FIND("(",#{my_cell}))),NOT(ISERROR(FIND(")",#{my_cell}))),NOT(ISERROR(FIND(" ",#{my_cell}))),NOT(ISERROR(FIND("+",#{my_cell}))),NOT(ISERROR(FIND("*",#{my_cell}))),NOT(ISERROR(FIND("-",#{my_cell}))))}
      },
      {
        :key => :range_validation,
        :formula => %q{AND(NOT(ISBLANK(#{my_cell})), #{range_of_validation.map{|cell| 'NOT(UPPER(#{my_cell})=#{cell})'}.join(',')}}
      },
      {
        :key => :range_validation_but_better,
        :formula => %q{AND(NOT(ISBLANK(#{my_cell})),NOT(OR(EXACT(UPPER(#{my_cell}),#{range_of_validation}))))}
      }
    ]
  end

  def styles
    [
      {
        :key => :red_when_incomplete_entry,
        :rule_key => :incomplete_entry,
        :style_definition => { :fg_color => "428751", :type => :dxf}
      },
      {
        :key => :red_when_incorrect_entry,
        :rule_key => :incorrect_entry,
        :style_definition => { :fg_color => "428751", :type => :dxf }
      }
    ]
  end

  def conditional_formattings
    [{
      :style_key => :red_when_incomplete_entry,
      :labels => ["Lysed?", "HMDMC"]
    }]
  end

  def manifest_templates
    [{
      :name => 'full_tube_manifest',
      :headers =>  %Q{Tube Barcode
Sanger Barcode
SANGER SAMPLE ID
HMDMC
DNA SOURCE
SUPPLIER SAMPLE NAME
Donor ID
DATE OF SAMPLE COLLECTION (yyyy-mm-dd only)
SAMPLE TYPE
Lysed?
VOLUME (ul)
GENDER
IS SAMPLE A CONTROL?
IS RE-SUBMITTED SAMPLE?
STORAGE CONDITIONS
MOTHER (optional)
FATHER (optional)
SIBLING (optional)
GC CONTENT
PUBLIC NAME
TAXON ID
SCIENTIFIC NAME
SAMPLE ACCESSION NUMBER (optional)}.split("\n")
    },{
      :name => 'half_tube_manifest',
      :headers => %Q{Tube Barcode
Sanger Barcode
SANGER SAMPLE ID
HMDMC
DNA SOURCE}.split("\n")
    }]
  end

  def headers(manifest_template_name)
    manifest_templates.select{|t| t[:name] == manifest_template_name}.first[:headers]
  end

  def data_validations
    [
      {
        :type => :list,
        :valid_values => ["Test1", "Test2", "Test3"],

        :showDropDown => false,
        :showErrorMessage => true,
        :errorTitle => '',

        :errorStyle => :stop,
        :showInputMessage => true,
        :promptTitle => '',
        :labels => ['DNA SOURCE']
      },
      {
        :valid_values => ["1234", "Prueba2", "Prueba3", "prueba4"],
        :type => :list,
        :labels => %Q{Tube Barcode
Sanger Barcode
SANGER SAMPLE ID
HMDMC
SUPPLIER SAMPLE NAME
Donor ID
DATE OF SAMPLE COLLECTION (yyyy-mm-dd only)
SAMPLE TYPE
Lysed?
VOLUME (ul)
GENDER
IS SAMPLE A CONTROL?
IS RE-SUBMITTED SAMPLE?
STORAGE CONDITIONS
MOTHER (optional)
FATHER (optional)
SIBLING (optional)
GC CONTENT
PUBLIC NAME
TAXON ID
SCIENTIFIC NAME
SAMPLE ACCESSION NUMBER (optional)}.split("\n")
        }
    ]
  end

def label_definitions
  @definitions
end

  def label_definition(header)
    label_definitions.select{|d| d[:labels].include?(header)}.first
  end

  def column_headers
    label_definitions.map{|d| d[:labels]}.flatten.uniq.compact
  end

  def data_validation_options(header)
    label_definition(header)[:data_validation]
  end

  def each_label_definition(&block)
    label_definitions.each do |l|
      yield l
    end
  end

  def definitions_with_data_validation(&block)
    label_definitions.select{|l| l[:data_validation][:type] == :list}
  end

end
