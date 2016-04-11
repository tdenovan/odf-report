module ODFReport

  class Row
    include Nested

    def initialize(opts)
      @name             = opts[:name]
      @collection_field = opts[:collection_field]
      @collection       = opts[:collection]

      @fields = []
      @texts = []
      @tables = []
      @sections = []

    end

    def replace!(doc, row = nil)

      return unless @row_node = find_row_node(doc)

      @collection = get_collection_from_item(row, @collection_field) if row

      @collection.each do |data_item|

        new_row = get_row_node

        @tables.each    { |t| t.replace!(new_row, data_item) }

        @sections.each  { |s| s.replace!(new_row, data_item) }

        @texts.each     { |t| t.replace!(new_row, data_item) }

        @fields.each    { |f| f.replace!(new_row, data_item) }

        @row_node.before(new_row)

      end

      @row_node.remove

    end # replace_row

    def remove!(doc)
      return unless row = find_row_node(doc)

      row.remove
    end # replace_row

  private

    def find_row_node(doc)
      rows = doc.xpath("//w:tr[w:tc[w:tbl[w:tblPr[w:tblCaption[@w:val='#{@name}']]]]]")

      rows.empty? ? nil : rows.first

    end

    def get_row_node
      node = @row_node.dup

      name = node.get_attribute('text:name').to_s
      @idx ||=0; @idx +=1
      node.set_attribute('text:name', "#{name}_#{@idx}")

      node
    end

  end

end
