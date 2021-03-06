module ODFReport

class Report

  FILES_TO_UPDATE = {
    doc: [/document.xml.rels/, /document.xml/, /drawing/, /chart/],
    ppt: [/presentation.xml.rels/, /slides\/slide.*\.xml$/, /drawing/, /presentation.xml$/, /\[Content_Types\]/],
    excel: ['xl/tables/table1.xml', 'xl/worksheets/sheet1.xml', 'xl/sharedStrings.xml']
  }

  def initialize(template_name, &block)

    @file = ODFReport::File.new(template_name)

    case ::File::extname(template_name)
      when '.docx' then @file_type = :doc
      when '.pptx' then @file_type = :ppt
      when '.xlsx' then @file_type = :excel
    end

    @texts = []
    @fields = []
    @text_fields = []
    @tables = []
    @sections = []
    @slides = []
    @remove_tables = []
    @remove_sections = []
    @remove_slides = []
    @charts = []
    @spreadsheets = []
    @remove_rows = []

    # Manager singleton classes
    @relationship_manager = ODFReport::RelationshipManager.new(@file, @file_type)
    @image_manager = ODFReport::ImageManager.new(@relationship_manager, @file)
    @slide_manager = ODFReport::SlideManager.new(@relationship_manager, @file)
    @chart_manager = ODFReport::ChartManager.new(@relationship_manager, @file)
    @table_manager = ODFReport::TableManager.new(@file_type)

    yield(self)

  end

  def add_field(field_tag, value='')
    opts = {:name => field_tag, :value => value}
    field = Field.new(opts)
    @fields << field
    @table_manager.add_variables(field_tag, value)
  end

  alias_method :add_header, :add_field

  def add_text(field_tag, value='')
    opts = {:name => field_tag, :value => value}
    text = Text.new(opts)
    @texts << text
  end

  def add_text_field(field_tag, value='')
    opts = {:name => field_tag, :value => value}
    text_field = TextField.new(opts)
    @text_fields << text_field
    @table_manager.add_variables(field_tag, value)
  end

  def add_table(table_name, collection, opts={})
    opts.merge!(:name => table_name, :collection => collection, :file_type => @file_type)
    tab = Table.new(opts, @image_manager)
    @tables << tab

    yield(tab)
  end

  def remove_table(table_name, collection, opts={})
    opts.merge!(:name => table_name, :collection => collection, :file_type => @file_type)
    tab = Table.new(opts, @image_manager)
    @remove_tables << tab
  end

  def add_chart(chart_name, collection, opts={})
    opts.merge!(:name => chart_name, :collection => collection, :file => @file)
    chart = Chart.new(opts)
    @charts << chart
    @chart_manager.add_charts(chart_name, collection, opts={})
  end

  def add_spreadsheet(spreadsheet_name, collection, opts={})
    opts.merge!(:name => spreadsheet_name, :collection => collection)
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

  def remove_row(row_name, opts={})
    opts.merge!(:name => row_name)
    row = Row.new(opts)
    @remove_rows << row
  end

  # <<<<<<<<<<<<<<<<<<<
  # Changes by tdenovan
  # <<<<<<<<<<<<<<<<<<<

  # Added by tdenovan
  def add_slide(slide_name, collection, insert_before_slide, opts={})
    opts.merge!(name: slide_name, collection: collection, insert_before_slide: insert_before_slide)
    slide = Slide.new(opts, @slide_manager)
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

      file.update_files(*FILES_TO_UPDATE[@file_type]) do |txt, filename|
        puts filename

        parse_document(txt) do |doc|
          @relationship_manager.parse_relationships(doc) if filename == RelationshipManager::RELATIONSHIP_FILES[:ppt] or filename == RelationshipManager::RELATIONSHIP_FILES[:doc]
          @image_manager.find_image_ids(doc)
          @chart_manager.find_chart_ids(doc, filename) # Scan each chart

          @slides.each         { |s| s.replace!(doc, nil, filename) } if /slides\/slide.*\.xml$/ =~ filename
          @sections.each       { |s| s.replace!(doc) }
          @tables.each         { |t| t.replace!(doc) }
          @texts.each          { |t| t.replace!(doc) }
          @fields.each         { |f| f.replace!(doc) }
          @text_fields.each    { |f| f.replace!(doc, filename) }
          @charts.each         { |c| c.replace!(doc, filename) }
          @spreadsheets.each   { |c| c.replace!(doc, filename) } if @file_type == :excel # Extract chart from docx zip and store it locally

          # CHANGED by tdenovan. We need a special call to remove template slides, as it can't be done in the replace method as other slides might rely on the template slide
          @slide_manager.update_presentation_file doc if /presentation.xml$/ =~ filename
          @slide_manager.update_content_type_file doc if /\[Content_Types\]/ =~ filename
          @remove_tables.each  { |t| t.remove!(doc) }
          @remove_rows.each { |r| r.remove!(doc) } if 'word/document.xml' === filename

          @table_manager.validate_row(doc, filename)
        end

      end

      @relationship_manager.write_new_relationships if @file_type == :doc or @file_type == :ppt
      @image_manager.write_images
      @slide_manager.write_slides
      @chart_manager.write_charts # Modify xlsx and save altered, delete old file

      # Notify chart something of the altered file...

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
