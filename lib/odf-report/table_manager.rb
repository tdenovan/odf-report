require 'dentaku'
module ODFReport

  class TableManager

    def validate_row(doc, filename)

      return unless /word\/document/ === filename

      calculator = Dentaku::Calculator.new
      doc.xpath("//w:t").each do |txt| # Go through every <w:t>
        txt.inner_html.scan(/\{\{.*\((.*)\)\}\}/).each do |arg| # Scan to see if there's a condition
          condition = $1.gsub(/&gt;/, '>').gsub(/&lt;/, '<') # Convert < and > signs
          if calculator.evaluate condition # Evaluate Condition
            txt.xpath("ancestor::*[w:tc]").remove # Remove Row
          else
            txt.inner_html = txt.inner_html.gsub(/\s*\{\{.*\}\}/, '') # Remove Condition
          end
        end
      end

    end

  end # TableManager class
end # ODFReport module