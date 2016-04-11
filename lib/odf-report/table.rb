module ODFReport

class Table
  include Nested

  def initialize(opts, image_manager)
    @name             = opts[:name]
    @collection_field = opts[:collection_field]
    @collection       = opts[:collection]
    @image_manager    = image_manager
    @file_type = opts[:file_type]
    case @file_type
      when :doc then @name_space = 'w'
      when :ppt then @name_space = 'a'
    end

    @fields = []
    @texts = []
    @tables = []
    @images = []
    @charts = []

    @template_rows = []
    @header           = opts[:header].nil? ? true : opts[:header]
    @skip_if_empty    = opts[:skip_if_empty] || false
  end

  def replace!(doc, row = nil)

    if @file_type == :doc and doc.namespaces.include? 'xmlns:w' and doc.xpath("//w:document").any? or# Look through document.xml
       @file_type == :ppt and doc.namespaces.include? 'xmlns:a'
      return unless table = find_table_node(doc)

      @template_rows = table.xpath(".//#{@name_space}:tr")

      # Disabling the below because it creates an odd effect on the table
      # @header = table.xpath("//w:tblHeader").empty? ? @header : false

      @collection = get_collection_from_item(row, @collection_field) if row
      if @skip_if_empty && @collection.empty?
        table.remove
        return
      end
      @collection.each do |data_item|
        new_node = get_next_row
        @tables.each    { |t| t.replace!(new_node, data_item) }
        @images.each    { |i| @image_manager.register_new_image(i.dup, new_node, data_item) }
        @texts.each     { |t| t.replace!(new_node, data_item) }
        @fields.each    { |f| f.replace!(new_node, data_item) }
        @charts.each    { |c| c.replace!(new_node, data_item) }
        table.add_child(new_node)

      end
      @template_rows.each_with_index do |r, i|
        r.remove if (get_start_node..template_length) === i
      end

    end

  end # replace


  def remove!(doc)

    if doc.namespaces.include? 'xmlns:w' and doc.xpath("//w:document").any?
      return unless table = find_table_node(doc)

      table.remove
    end
  end # replace_section

  def remove_row!(row_text_content)
    # TODO find the row containing that text content
  end

private

  def get_next_row
    @row_cursor = get_start_node unless defined?(@row_cursor)

    ret = @template_rows[@row_cursor]
    if @template_rows.size == @row_cursor + 1
      @row_cursor = get_start_node
    else
      @row_cursor += 1
    end
    return ret.dup
  end

  def get_start_node
    @header ? 1 : 0
  end

  def template_length
    @tl ||= @template_rows.size
  end

  def find_table_node(doc)
    case @file_type
      when :doc then tables = doc.xpath("//w:tbl[w:tblPr[w:tblCaption[@w:val='#{@name}']]]")
      when :ppt then tables = doc.xpath("//p:graphicFrame[p:nvGraphicFramePr[p:cNvPr[@title='#{@name}']]]/a:graphic/a:graphicData/a:tbl")
    end
    tables.empty? ? nil : tables.first
  end

end

end
