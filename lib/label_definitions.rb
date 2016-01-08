require_relative 'excel_manifest'

class LabelDefinitions

  include ExcelManifest

  def rules
    [
      {
        :name => :incorrect_format,
        :formula => %q{OR(LEN(#{my_cell})>15,NOT(ISERROR(FIND("&amp;",#{my_cell}))),NOT(ISERROR(FIND("(",#{my_cell}))),NOT(ISERROR(FIND(")",#{my_cell}))),NOT(ISERROR(FIND(" ",#{my_cell}))),NOT(ISERROR(FIND("+",#{my_cell}))),NOT(ISERROR(FIND("*",#{my_cell}))),NOT(ISERROR(FIND("-",#{my_cell}))))}
      },
      {
        :name => :incorrect_content,
        :formula => %q{OR(NOT(ISERROR(FIND(".",#{my_cell}))),NOT(ISERROR(FIND(",",#{my_cell}))),NOT(ISERROR(FIND("\\",#{my_cell}))),NOT(ISERROR(FIND("\/",#{my_cell}))))}
      },
      {
        :name => :empty_entry,
        :formula => %q{AND(ISBLANK(#{my_cell}),COUNTA(#{range_of_entry})&gt;0)}
      },
      {
        :name => :date_validation,
        :formula => %q{OR(AND(NOT(ISBLANK(#{my_cell})),NOT(LEN(#{my_cell})=5),NOT(ISERROR(FIND("/",#{my_cell})))),AND(NOT(ISBLANK(#{my_cell})),NOT(LEN(#{my_cell})=4),(ISERROR(FIND("/",#{my_cell})))),NOT(ISERROR(FIND(" ",#{my_cell}))))}
      },
      {
        :name => :number_validation,
        :formula => %q{AND(NOT(ISNUMBER(#{my_cell})),NOT(ISBLANK(#{my_cell})))}
      },
      {
        :name => :string_validation_30,
        :formula => %q{OR(LEN(#{my_cell})&gt;30,NOT(ISERROR(FIND("&amp;",#{my_cell}))),NOT(ISERROR(FIND("(",#{my_cell}))),NOT(ISERROR(FIND(")",#{my_cell}))),NOT(ISERROR(FIND("$",#{my_cell}))),NOT(ISERROR(FIND("*",#{my_cell}))),NOT(ISERROR(FIND("-",#{my_cell}))))"}
      },
      {
        :name => :string_validation_7,
        :formula => %q{OR(LEN(#{my_cell})&gt;7,NOT(ISERROR(FIND("&amp;",#{my_cell}))),NOT(ISERROR(FIND("(",#{my_cell}))),NOT(ISERROR(FIND(")",#{my_cell}))),NOT(ISERROR(FIND(" ",#{my_cell}))),NOT(ISERROR(FIND("+",#{my_cell}))),NOT(ISERROR(FIND("*",#{my_cell}))),NOT(ISERROR(FIND("-",#{my_cell}))))}
      },
      {
        :name => :range_validation,
        :formula => %q{AND(NOT(ISBLANK(#{my_cell})), #{range_of_validation.map{|cell| 'NOT(UPPER(#{my_cell})=#{cell})'}.join(',')}}
      },
      {
        :name => :range_validation_but_better,
        :formula => %q{AND(NOT(ISBLANK(#{my_cell})),NOT(OR(EXACT(UPPER(#{my_cell}),#{range_of_validation}))))}
      }
    ]
  end

  def styles
    [
      {
        :name => :incomplete_entry,
        :fg_color => "428751",
        :type => :dxf
      },
      {
        :name => :incorrect_entry,
        :fg_color => "428751", :type => :dxf
      }
    ]
  end

  def conditional_formattings
    [{
      :rule => :number_validation,
      :style => :incorrect_entry,
      :labels => ["Lysed?", "HMDMC"]
      }]
  end

def initialize
  @definitions =   [{
    :data_validation => {
      :type => :list,

      :valid_values => ["Test1", "Test2", "Test3"],
      #:formula1 => 'C1:C8',
      #:error => 'Only values from C1:C8',
      #:prompt => 'Only values from C1:C8'

      :showDropDown => false,
      :showErrorMessage => true,
      :errorTitle => '',

      :errorStyle => :stop,
      :showInputMessage => true,
      :promptTitle => '',

      },
    :conditional_formatting => {

    },
    :labels => ['DNA SOURCE']
  },
    {
     :data_validation => {
       :valid_values => ["1234", "Prueba2", "Prueba3", "prueba4"],
      :type => :list,

     },
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
