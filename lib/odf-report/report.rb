module ODFReport

class Report

  FILES_TO_UPDATE = {
    doc: ['word/document.xml', /chart/, RelationshipManager::RELATIONSHIP_FILE],
    excel: ['xl/tables/table1.xml', 'xl/worksheets/sheet1.xml', 'xl/sharedStrings.xml']
  }

  def initialize(template_name, &block)

    @file = ODFReport::File.new(template_name)
    @file_type = :doc

    case ::File::extname(template_name)
    when '.xlsx'
      @file_type = :excel
    end

    @texts = []
    @fields = []
    @tables = []
    @sections = []
    @slides = []
    @remove_sections = []
    @remove_slides = []
    @charts = []
    @spreadsheets = []

    # Manager singleton classes
    @relationship_manager = ODFReport::RelationshipManager.new(@file)
    @image_manager = ODFReport::ImageManager.new(@relationship_manager, @file)

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
    tab = Table.new(opts, @image_manager)
    @tables << tab

    yield(tab)
  end

  def add_chart(chart_name, collection, opts={})
    opts.merge!(:name => chart_name, :collection => collection)
    chart = Chart.new(opts)
    @charts << chart
    # opts.merge!(:name => chart_name, :collection => collection)
    spreadsheet = Spreadsheet.new(opts)
    @spreadsheets << spreadsheet
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
    @image_manager.add_existing_image(name, path)
  end

  def generate(dest = nil)

    @file.update_content do |file|

      # file.update_files(FILES_TO_UPDATE[@file_type]) do |txt, filename|
      # file.update_files('word/document.xml', /chart/, RelationshipManager::RELATIONSHIP_FILE) do |txt, filename|
      file.update_files('xl/tables/table1.xml', 'xl/worksheets/sheet1.xml', 'xl/sharedStrings.xml') do |txt, filename|

        parse_document(txt) do |doc|
          @relationship_manager.parse_relationships(doc) if filename == RelationshipManager::RELATIONSHIP_FILE
          @image_manager.find_image_ids(doc)

          @slides.each         { |s| s.replace!(doc) }
          @sections.each       { |s| s.replace!(doc) }
          @tables.each         { |t| t.replace!(doc) }
          @texts.each          { |t| t.replace!(doc) }
          @fields.each         { |f| f.replace!(doc) }
          @charts.each         { |c| c.replace!(doc, filename) }
          @spreadsheets.each   { |c| c.replace!(doc, filename) } if @file_type == :excel

          # CHANGED by tdenovan. We need a special call to remove template slides, as it can't be done in the replace method as other slides might rely on the template slide
          @slides.each  { |s| s.remove!(doc) }

        end

      end

      @relationship_manager.write_new_relationships if @file_type == :doc
      @image_manager.write_images
    end

    # Write the docx file
    ::File.open(dest, "wb") {|f| f.write(@file.data) }

    # delete_old_images(@file, dest)

  end

private

  def parse_document(txt)
    doc = Nokogiri::XML(txt)
    yield doc
    txt.replace(doc.to_xml(:save_with => Nokogiri::XML::Node::SaveOptions::AS_XML))

  end

end

end
