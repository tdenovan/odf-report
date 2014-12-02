require 'dentaku'
module ODFReport

  class TableManager

    def initialize
      @calculator = Dentaku::Calculator.new
    end

    def add_variables(name, value)
      return unless value.is_a? Fixnum or value.is_a? Float
      @calculator.store(name => value)
    end

    def validate_row(doc, filename)

      return unless filename == 'word/document.xml' # Go through word/document.xml

      doc.xpath("//w:tr").each do |row|
        row.xpath("descendant::*[w:t]//w:t").inner_html.scan(/\{\{.*\((.*)\)\}\}/).each do |arg| # Scan to see if there's a condition
          condition = $1.gsub(/&gt;/, '>').gsub(/&lt;/, '<') # Convert < and > signs
          if @calculator.evaluate condition # Evaluate Condition
            row.xpath("*//w:t").each { |txt| txt.inner_html = txt.inner_html.gsub(/\{\{.*\}\}/, '').gsub(/\{\{.*/, '').gsub(/\}\}/, '') } # Remove condition
          else
            row.remove # Remove Row
          end
        end
      end

    end

  end # TableManager class
end # ODFReport module