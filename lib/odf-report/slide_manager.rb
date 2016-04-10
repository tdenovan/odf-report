module ODFReport

  # This class is responsible for managing the list of slides in the presentation that will be replaced
  # It also keeps a record of new slides that need to be written
  # This class is effectively a singleton
  class SlideManager
    attr_accessor :existing_slides, :new_slides

    SLIDE_DIR_NAME = "ppt/slides"

    def initialize(relationship_manager, file_manager)
      @relationship_manager = relationship_manager
      @slides = []
      @slide_rels = [] # relationship files
      @file_manager = file_manager
      @tag_idx = 10000
    end

    # Creates a new slide by copying an old one
    # @param slide_path is the existing slide_path that should be duplicated
    # @slide_number is the current slide number
    def duplicate_slide(slide_path, slide_number)
      old_slide_filename = ::File.basename slide_path
      old_slide_dir = ::File.dirname slide_path
      new_slide_filename = "slide#{slide_number}.xml"
      new_slide_rels_filename = "#{new_slide_filename}.rels"
      
      # Duplicate the slide 
      slide = {is_new: true, path: new_slide_filename, old_path: slide_path}
      slide[:id], slide[:target] = @relationship_manager.new_relationship(:slide, "slides/#{slide[:path]}")
      slide[:doc] = Nokogiri::XML(@file_manager.read_file(slide_path))
      slide[:number] = slide_number
      create_new_slide_id(slide[:doc])
      @slides << slide
      
      # Duplicate the relationship file for the slide
      rel_file = {old_path: "#{old_slide_dir}/_rels/#{old_slide_filename}.rels", path: "#{::File.basename(slide[:target])}.rels" }
      rel_file[:doc] = Nokogiri::XML(@file_manager.read_file(rel_file[:old_path]))
      parse_slide_relationship_file rel_file[:doc], rel_file
      @slide_rels << rel_file
      
      # Return the xml node for the slide
      return slide[:doc].xpath(".//p:cSld").first
    end
    
    # Creates a new unique id for a duplicated slide. Slides can't have the same id
    def create_new_slide_id(doc)
      node = doc.xpath(".//p14:creationId", { p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main'}).first
      node.set_attribute('val', Random.rand(999999000))
    end
    
    # Updates the presenation.xml file
    # Writes any new slides to the presentation.xml file
    # Deletes any template slides (e.g. slides that were duplicated) from the presentation.xml file
    def update_presentation_file(doc)
      # Write the new slides to presentation.xml
      @slides.reverse.each do |slide|
        rel_id = @relationship_manager.get_relationship_by_target("slides/#{::File.basename(slide[:old_path])}")[:id]
        old_slide_node = doc.xpath(".//p:sldId[@r:id='#{rel_id}']").first
        slide_node = old_slide_node.dup
        slide_node.set_attribute('id', slide[:number])
        slide_node.set_attribute('r:id', slide[:id])
        old_slide_node.after slide_node
      end
      
      # Delete the old template slides from presentation.xml
      template_slide_names = @slides.map { |slide| slide[:old_path] }.uniq
      template_slide_names.each do |template_slide_name|
        rel_id = @relationship_manager.get_relationship_by_target("slides/#{::File.basename(template_slide_name)}")[:id]
        doc.xpath(".//p:sldId[@r:id='#{rel_id}']").remove
      end
    end
    
    # Parse the slide relationship doc
    def parse_slide_relationship_file(doc, slide_rel)
      slide_rel[:tags] = []
            
      # Duplicate and update references to any tags. tag.xml files can only be used in a single slide
      doc.xpath(".//xmlns:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags']").each do |tag|
        @tag_idx += 1
        slide_rel[:tags] << { old_path: tag.get_attribute('Target').to_s, new_filename: "tag#{@tag_idx}.xml" }
      end
    end
    
    # Updates the [CONTENT_TYPE].xml file to accommodate the new slides
    def update_content_type_file(doc)
      @slides.each do |slide|
        node = doc.xpath(".//xmlns:Override[@PartName='/ppt/slides/#{::File.basename(slide[:old_path])}']").first
        new_node = node.dup
        new_node.set_attribute('PartName', "/ppt/#{slide[:target]}")
        node.after new_node
      end
      
      @slide_rels.each do |rel_file|
        rel_file[:tags].each do |tag|
          # Create a new record in the [CONTENT_TYPES] xml file
          tag_filename = ::File.basename(tag[:old_path])
          node = doc.xpath(".//xmlns:Override[@PartName='/ppt/tags/#{tag_filename}']").first
          new_node = node.dup
          new_node.set_attribute('PartName', "/ppt/tags/#{tag[:new_filename]}")
          node.after new_node
        end
      end
    end
    
    # Cleans up a slide relationship file (e.g. deletes notes)
    def clean_up_slide_relationship_file(doc, rel_file)
      doc.xpath(".//xmlns:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide']").remove
            
      # Update tag references
      rel_file[:tags].each do |tag|
        node = doc.xpath(".//xmlns:Relationship[@Target='#{tag[:old_path]}']").first
        node.set_attribute 'Target', "../tags/#{tag[:new_filename]}"
      end
    end

    # Writes slides to the zip file
    def write_slides
      # Enumerate all new slides and write them to the zip file
      @slides.select { |slide| slide[:is_new] }.each do |new_slide|
        @file_manager.output_stream.put_next_entry("ppt/#{new_slide[:target]}")
        @file_manager.output_stream.write new_slide[:doc].to_xml(:save_with => Nokogiri::XML::Node::SaveOptions::AS_XML)
      end
               
      # Enumerate all relationship files and write them out
      @slide_rels.each do |rel_file|
        clean_up_slide_relationship_file rel_file[:doc], rel_file
        @file_manager.output_stream.put_next_entry("ppt/slides/_rels/#{rel_file[:path]}")
        @file_manager.output_stream.write rel_file[:doc].to_xml(:save_with => Nokogiri::XML::Node::SaveOptions::AS_XML)
        
        # Duplicate the tags
        rel_file[:tags].each do |tag|
          tag_filename = ::File.basename(tag[:old_path])
          @file_manager.output_stream.put_next_entry "ppt/tags/#{tag[:new_filename]}"
          @file_manager.output_stream.write @file_manager.read_file("ppt/tags/#{tag_filename}")
        end
      end
    end

  end # ImageManager class
end # ODFReport module
