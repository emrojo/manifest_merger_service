module ExcelManifest
  EXCEL_COLUMNS = ('A'..'Z').to_a
  DEFAULT_ROW_RANGE = (1..99)

  def index_to_column_name(value)
    return EXCEL_COLUMNS[value] if value < EXCEL_COLUMNS.length
    index_column_name_list = []
    while value >= EXCEL_COLUMNS.length do
      index_column_name_list.push(value % EXCEL_COLUMNS.length)
      value = value / EXCEL_COLUMNS.length
    end
    index_column_name_list.push(value)
    index_column_name_list.reverse.map{|c| EXCEL_COLUMNS[c-1]}.join('')
  end

  def column_selection_to_str(column_name, start, size)
    "#{column_name}#{start+1}:#{column_name}#{size+start+1}"
  end

  def update_table_with_table(tableA, tableB, startARow, startACol)
    tableB.each_with_index do |row, idx|
      tableA[startARow + idx] = tableA[startARow + idx].slice(0, startACol).concat(row)
    end
  end

  def update_list_with_list(listA, listB, startCol)
    listA = listA.slice(0, startACol).concat(listB)
  end

  def fill_sheet_with_empty_rows_until_row_pos(sheet, row_pos)
    (row_pos - sheet.rows.count).times { sheet.add_row [] }
  end

  def update_sheet_with_table(sheet, table, startARow, startACol)
    fill_sheet_with_empty_rows_until_row_pos(sheet, startARow)
    table.each_with_index do |row, idx|
      if sheet.rows[startARow + idx].nil?
        sheet.add_row([].fill(nil, 0, startACol-1).concat(row))
      else
        update_list_with_list(sheet.rows[startARow + idx], table[idx], startACol)
      end
    end
  end

  def transpose_table(table)
    # NB: transposed_table is made to work with arrays of different lengths so it could not be transposed
    # with a simple :transpose call
    max_size = table.map(&:count).max
    table.map {|row| row.fill(nil, row.length, (max_size - row.length))}.transpose
  end
end
