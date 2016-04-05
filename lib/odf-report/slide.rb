module ODFReport

  class Slide
    include Nested
    @@idx = 0

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

    def replace!(doc, row = nil)
      binding.pry
      return unless @slide_node = find_slide_node(doc)
      binding.pry
      # TODO get anchor (i.e. slide that this one will be inserted prior to)

      @collection = get_collection_from_item(row, @collection_field) if row

      @collection.each do |data_item|

        new_slide = get_slide_node

        @tables.each    { |t| t.replace!(new_slide, data_item) }
        @sections.each  { |s| s.replace!(new_slide, data_item) }
        @texts.each     { |t| t.replace!(new_slide, data_item) }
        @fields.each    { |f| f.replace!(new_slide, data_item) }

        @slide_node.before(new_slide)

      end

      #CHANGED note the template node cannot be deleted here as it may be premature. If there are multiple calls to add_slide which rely on the same template, deleting it here would cause issues

    end # replace_slide

    def remove!(doc)
      return unless slide = find_slide_node(doc)

      slide.remove
    end # replace_slide

  private

    def find_slide_node(doc)

      slides = doc.xpath(".//xmlns:cSld[@name='#{@name}']")

      slides.empty? ? nil : slides.first

    end

    def get_slide_node
      node = @slide_node.dup

      name = node.get_attribute('text:name').to_s
      @@idx ||=0; @@idx +=1
      node.set_attribute('text:name', "#{name}_#{@idx}")

      node
    end

  end

end
