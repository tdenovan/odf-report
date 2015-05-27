module ODFReport

  # Class for creating images (e.g. on tables, inside a doc, etc)
  class Image
    attr_accessor :name, :path, :id, :target, :is_new

    DELIMITERS = %w([ ])

    def initialize(name, path, data_field = nil, is_new = false)
      @name = name
      @path = path
      @id = nil
      @target = nil
      @is_new = is_new
      @data_field = data_field # data field is only used if this image is part of a collection (e.g. on a table) where the value will be populated later. If data_field is populated, path should be nil
    end

    # Data item path
    def set_path(data_item)
      @path = data_item[@data_field] unless data_item.nil? or @data_field.nil?
    end

    def replace_new_image_id!(content, data_item = nil)
      current_id = nil

      # replace the id of the image
      if content.namespaces.include? 'xmlns:w' and content.xpath("//w:drawing").any? # Looking through word/document.xml
        img = content.xpath("//w:drawing//wp:docPr[@title='#{@name}']/following-sibling::*").xpath(".//a:blip", {'a' => "http://schemas.openxmlformats.org/drawingml/2006/main"}).first
        current_id = img.attributes['embed'].value
      end

      txt = content.inner_html
      txt.gsub!("r:embed=\"#{current_id}\"", "r:embed=\"#{@id}\"")
      content.inner_html = txt
    end # replace! method

  end
end
