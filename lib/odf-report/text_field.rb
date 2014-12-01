require 'debugger'

module ODFReport
  class TextField

    def initialize(opts, &block)
      @name = opts[:name]
      @value = opts[:value] || ''
    end

    def replace!(doc, filename, data_item = nil)

      return unless /word\/document/ === filename

      @name = @name.to_s
      @value = @value.to_s

      doc.xpath("//w:default[@w:val='#{@name}']/ancestor::*[w:fldChar]/following-sibling::*[w:t]//w:t").each {|m| m.inner_html = @value }
    end

  end
end
