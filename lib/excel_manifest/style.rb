class Style

  @@all = nil

  def self.all
    @@all
  end

  def self.all=(value)
    @@all=value
  end

  attr_accessor :key
  attr_accessor :rule
  attr_accessor :definition

  def initialize(params)
    @key = params[:key]
    @rule = params[:rule]
    @definition = params[:style_definition]
  end
end


