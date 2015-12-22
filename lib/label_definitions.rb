require_relative 'excel_manifest'

class LabelDefinitions

  include ExcelManifest

def initialize
  @definitions =   [{
    :type => :list,
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
     :labels => ['DNA SOURCE']
  },
    {
    :type => :string,
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
end
