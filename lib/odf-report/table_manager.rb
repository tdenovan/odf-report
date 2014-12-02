require 'dentaku'
module ODFReport

  class TableManager

    def initialize
      @calculator = Dentaku::Calculator.new
      @active = false
    end

    def add_variables(name, value)
      return if value.nil? or value.values.include? nil
      hash = {name => value}
      @active = true

      if value.is_a? Hash
        @calculator.store(value)
      else
        @calculator.store(hash)
      end
    end

    def validate_row(doc, filename)

      return unless /word\/document/ === filename and @active

      doc.xpath("//w:tbl//w:t").each do |txt| # Go through every table text
        txt.inner_html.scan(/\{\{.*\((.*)\)\}\}/).each do |arg| # Scan to see if there's a condition
          condition = $1.gsub(/&gt;/, '>').gsub(/&lt;/, '<') # Convert < and > signs
          if @calculator.evaluate condition # Evaluate Condition
            txt.inner_html = txt.inner_html.gsub(/\s*\{\{.*\}\}/, '') # Remove Condition
          elsif @calculator.evaluate(condition).nil?
          else
            txt.xpath("ancestor::*[w:tc]").remove # Remove Row
          end

        end
      end

    end

  end # TableManager class
end # ODFReport module