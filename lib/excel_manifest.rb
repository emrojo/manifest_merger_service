module ExcelManifest
  EXCEL_COLUMNS = ('A'..'Z').to_a
  DEFAULT_ROW_RANGE = (1..99)

  def index_to_column_name(value)
    index_column_name_list = []
    while value >= EXCEL_COLUMNS.length do
      index_column_name_list.push(value % EXCEL_COLUMNS.length)
      value = value / EXCEL_COLUMNS.length
    end
    index_column_name_list.push(value)
    index_column_name_list.reverse.map{|c| EXCEL_COLUMNS[c]}.join('')
  end

  def column_selection_to_str(column_name, start, size)
    "#{column_name}#{start}:#{column_name}#{size-start + 1}"
  end

end
