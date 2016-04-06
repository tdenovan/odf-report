module ODFReport

  class Slide
    include Nested
    @@idx = 1000

    def initialize(opts, slide_manager)
      @name             = opts[:name]
      @collection_field = opts[:collection_field]
      @collection       = opts[:collection]
      @insert_before_slide = opts[:insert_before_slide]
      @slide_manager = slide_manager

      @fields = []
      @texts = []
      @tables = []
      @sections = []

    end

    def replace!(doc, row = nil, slide_path = nil)
      
      return unless @slide_node = find_slide_node(doc)
      
      # TODO get anchor (i.e. slide that this one will be inserted prior to)

      @collection = get_collection_from_item(row, @collection_field) if row

      new_slide = get_slide_node(slide_path)

      @tables.each    { |t| t.replace!(new_slide, @collection) }
      @sections.each  { |s| s.replace!(new_slide, @collection) }
      @texts.each     { |t| t.replace!(new_slide, @collection) }
      @fields.each    { |f| f.replace!(new_slide, @collection) }

      #CHANGED note the template node cannot be deleted here as it may be premature. If there are multiple calls to add_slide which rely on the same template, deleting it here would cause issues

    end # replace_slide

    def remove!(doc)
      return unless slide = find_slide_node(doc)

      slide.remove
    end # replace_slide

  private

    def find_slide_node(doc)
      slides = doc.xpath(".//p:cSld[@name='#{@name}']")
      slides.empty? ? nil : slides.first
    end

    # @param old_slide_path is the path of the 
    def get_slide_node(old_slide_path)
      @@idx +=1
      
      node = @slide_manager.duplicate_slide old_slide_path, @@idx

      name = node.get_attribute('name').to_s
      node.set_attribute('name', "#{name}_#{@@idx}")

      node
    end

  end

end
