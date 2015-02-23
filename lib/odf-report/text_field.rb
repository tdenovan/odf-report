require 'debugger'

module ODFReport
  class TextField

    def initialize(opts, &block)
      @name = opts[:name]
      @value = opts[:value] || ''
    end

    def replace!(doc, filename, data_item = nil)

      return unless /word\/document/ === filename

      name = "[#{@name.to_s.upcase}]"
      value = @value.to_s.gsub('-', '')

      doc.xpath("//w:default[@w:val='#{name}']").each do |field|
        field.xpath("ancestor::*[w:fldChar]/following-sibling::*/w:t").first.inner_html = value
      end
    end
  end
end
