module ODFReport

class Report
  include Images

  def initialize(template_name, &block)

    @file = ODFReport::File.new(template_name)

    @texts = []
    @fields = []
    @tables = []
    @images = {}
    @image_names_replacements = {}
    @sections = []
    @slides = []
    @remove_sections = []
    @remove_slides = []

    yield(self)

  end

  def add_field(field_tag, value='')
    opts = {:name => field_tag, :value => value}
    field = Field.new(opts)
    @fields << field
  end

  def add_text(field_tag, value='')
    opts = {:name => field_tag, :value => value}
    text = Text.new(opts)
    @texts << text
  end

  def add_table(table_name, collection, opts={})
    opts.merge!(:name => table_name, :collection => collection)
    tab = Table.new(opts)
    @tables << tab

    yield(tab)
  end

  def add_section(section_name, collection, opts={})
    opts.merge!(:name => section_name, :collection => collection)
    sec = Section.new(opts)
    @sections << sec

    yield(sec)
  end

  def remove_section(section_name, opts={})
    opts.merge!(:name => section_name)
    sec = Section.new(opts)
    @remove_sections << sec
  end
  
  # <<<<<<<<<<<<<<<<<<<
  # Changes by tdenovan
  # <<<<<<<<<<<<<<<<<<<
  
  # Added by tdenovan
  def add_slide(slide_name, collection, insert_before_slide, opts={})
    opts.merge!(name: slide_name, collection: collection, insert_before_slide: insert_before_slide)
    slide = Slide.new(opts)
    @slides << slide
    
    yield(slide)
  end

  def remove_slide(slide_name, opts)
    opts.merge!(name: slide_name)
    slide = Slide.new(opts)
    @remove_slides << slide
  end
  
  # <<<<<<<<<<<<<<<<<<<<<<<
  # End changes by tdenovan
  # <<<<<<<<<<<<<<<<<<<<<<<

  def add_image(name, path)
    @images[name] = path
  end

  def generate(dest = nil)

    @file.update_content do |file|

      file.update_files('content.xml', 'styles.xml') do |txt|

        parse_document(txt) do |doc|

          @slides.each   { |s| s.replace!(doc) }
          @sections.each { |s| s.replace!(doc) }
          @tables.each   { |t| t.replace!(doc) }
          @texts.each    { |t| t.replace!(doc) }
          @fields.each   { |f| f.replace!(doc) }
        
          find_image_name_matches(doc)
          # avoid_duplicate_image_names(doc) # This method produces unreadable xml files for me
          
          # CHANGED by tdenovan. We need a special call to remove template slides, as it can't be done in the replace method as other slides might rely on the template slide
          @slides.each  { |s| s.remove!(doc) }

        end

      end

      replace_images(file)

    end

    if dest
      ::File.open(dest, "wb") {|f| f.write(@file.data) }
    else
      @file.data
    end

  end

private

  def parse_document(txt)
    doc = Nokogiri::XML(txt)
    yield doc
    txt.replace(doc.to_xml(:save_with => Nokogiri::XML::Node::SaveOptions::AS_XML))
  end

end

end
