class ConditionalFormatting
  attr_accessor :style
  attr_accessor :labels
  @@all = nil

  def self.all
    @@all
  end

  def self.all=(value)
    @@all=value
  end

  def initialize(params)
    @style = params[:style]
    @labels = params[:labels]
  end
end
