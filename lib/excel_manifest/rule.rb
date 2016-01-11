class Rule
  attr_accessor :key
  attr_writer :formula

  def initialize(params)
    @key = params[:key]
    @formula = params[:formula]
  end

end
