class Cell
  attr_accessor :row_position, :column_position

  def initialize(row, column)
    @row_position = row
    @column_position = column
  end
end
