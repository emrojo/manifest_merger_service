class WorkbookStyle
  attr_reader :style_id
  def initialize(workbook, style)
    @style_id = workbook.styles.add_style(style.definition)
  end

end
