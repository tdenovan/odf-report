require 'dentaku'
module ODFReport

  class TableManager

    def initialize(file_type)
      @calculator = Dentaku::Calculator.new
      @file_type = file_type
      case file_type
        when :doc then @name_space = 'w'
        when :ppt then @name_space = 'a'
      end
    end

    def add_variables(name, value)
      @calculator.store(name.to_s => value.to_s.gsub(/[^\d.-]/, '').to_f)
    end

    def validate_row(doc, filename)

      return unless filename == 'word/document.xml' or filename =~ /slide.*\.xml/ # Go through word/document.xml

      doc.xpath("//#{@name_space}:tr").each do |row|
        row.xpath("descendant::*[#{@name_space}:t]//#{@name_space}:t").inner_html.scan(/(\{\{.*\((.*)\)\}\})/).each do |arg| # Scan to see if there's a condition
          whole_condition = $1
          condition = $2.gsub(/&gt;/, '>').gsub(/&lt;/, '<') # Convert < and > signs
          if @calculator.evaluate condition # Evaluate Condition
            active = false
            row.xpath("*//#{@name_space}:t").each do |txt| # Remove condition
              active = true if txt.inner_html.include? "\{\{"
              next unless active
              inner_html = txt.inner_html.clone
              txt.inner_html = txt.inner_html.gsub(/\{\{.*\}\}/, '').gsub(/\{\{.*/, '').gsub(/.*\}\}/, '')
              txt.inner_html = '' if whole_condition.include? txt.inner_html
              break if inner_html.include? "\}\}"
            end
          else
            row.remove # Remove Row
          end
        end
      end

    end

  end # TableManager class
end # ODFReport module
