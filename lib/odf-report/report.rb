module ODFReport

class Report
  include Images

  def initialize(template_name, &block)

    @file = ODFReport::File.new(template_name)

    @texts = []
    @fields = []
    @tables = []
    @images = {}
    @sections = []
    @slides = []
    @remove_sections = []
    @remove_slides = []
    @charts = []

    # Image related variables
    @image_names_replacements = {}
    @image_name_id = {} # Creating a hash of image names and linking them with their id
    @image_id_paths = {}

    # Chart related variables
    $id_target = {}
    $name_id = {}

    yield(self)

  end

  def add_field(field_tag, value='')
    opts = {:name => field_tag, :value => value}
    field = Field.new(opts)
    @fields << field
  end

  alias_method :add_header, :add_field
  alias_method :add_title, :add_field
  alias_method :add_series, :add_field

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

  def add_chart(chart_name, collection, opts={})
    opts.merge!(:name => chart_name, :collection => collection)
    chart = Chart.new(opts)
    @charts << chart
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

      file.update_files('word/document.xml', /chart/, 'word/_rels/document.xml.rels') do |txt, filename|

        parse_document(txt) do |doc|

          @slides.each   { |s| s.replace!(doc) }
          @sections.each { |s| s.replace!(doc) }
          @tables.each   { |t| t.replace!(doc) }
          @texts.each    { |t| t.replace!(doc) }
          @fields.each   { |f| f.replace!(doc) }
          @charts.each   { |c| c.replace!(doc, filename) }

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
