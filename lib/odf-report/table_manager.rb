require 'dentaku'
module ODFReport

  class TableManager

    def initialize
      @calculator = Dentaku::Calculator.new
    end

    def add_variables(name, variable)
      instance_variable_set("@#{name}", variable)
      @calculator.store(instance_variable_get("@#{name}"))
    end

    def validate_row(doc, filename)

      return unless /word\/document/ === filename

      doc.xpath("//w:t").each do |txt| # Go through every <w:t>
        txt.inner_html.scan(/\{\{.*\((.*)\)\}\}/).each do |arg| # Scan to see if there's a condition
          condition = $1.gsub(/&gt;/, '>').gsub(/&lt;/, '<') # Convert < and > signs

          if @calculator.evaluate condition # Evaluate Condition
            txt.inner_html = txt.inner_html.gsub(/\s*\{\{.*\}\}/, '') # Remove Condition
          else
            txt.xpath("ancestor::*[w:tc]").remove # Remove Row
          end

        end
      end

    end

  end # TableManager class
end # ODFReport module