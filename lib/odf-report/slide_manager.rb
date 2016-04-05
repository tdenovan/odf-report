module ODFReport

  # This class is responsible for managing the list of slides in the presentation that will be replaced
  # It also keeps a record of new slides that need to be written
  # This class is effectively a singleton
  class SlideManager
    attr_accessor :existing_slides, :new_slides

    IMAGE_DIR_NAME = "word/media"

    def initialize(relationship_manager, file_manager)
      @relationship_manager = relationship_manager
      @images = []
      @file = file_manager
    end

    # Finds images ids of existing images in an xml doc (if any)
    def find_image_ids(doc)
      if doc.namespaces.include? 'xmlns:w' and doc.xpath("//w:drawing").any?
        @images.each do |image|

          doc.xpath("//w:drawing//wp:docPr[@title='#{image.name}']/following-sibling::*").xpath(".//a:blip", {'a' => "http://schemas.openxmlformats.org/drawingml/2006/main"}).each do |img|
            image.id = img.attributes['embed'].value
          end
        end # existing_images loop
      end
    end # find_image_ids method

    # Adds a reference to a template image that should be replaced
    def add_existing_image(name, path)
      @images << Image.new(name, path)
    end

    # Creates a new image - e.g. when a table row is duplicated
    def register_new_image(image, xml_node, data_item)
      image.set_path(data_item)
      image.is_new = true
      image.id, image.target = @relationship_manager.new_relationship(:image, image.path)
      image.replace_new_image_id!(xml_node, data_item)
      @images << image
      image
    end

    # Writes images to the zip file this includes images that are not being replaced, as the File class was modified so that not files in the word/media directory are written (in case the files is being replaced in this method and would be overwritten)
    # CHANGED - previously, all files in the word/media directory were written to the zip. Then, in this method, any images being replaced were overwritten. That approach caused the zip to have an invalid header. The revised approach is to write all files in the word/media directory once, and substitute any images being replaced at the time of the first (and only) writing
    def write_images
      # Enumerate all the files in the zip and write any that are in the media directory to the output buffer (which is used to generate the new zip file)
      @file.read_files do |entry| # entry is a file entry in the zip
        if entry.name.include? IMAGE_DIR_NAME
          # Check if this is an image being replaced
          current_image = @images.select { |image| !@relationship_manager.get_relationship(image.id).nil? and entry.name.include? @relationship_manager.get_relationship(image.id)[:target] }.first

          unless current_image.nil?
            replacement_path = current_image.path
            data = ::File.read(replacement_path)
          else
            entry.get_input_stream { |is| data = is.sysread }
          end

          @file.output_stream.put_next_entry(entry.name)
          @file.output_stream.write data
        end
      end

      # Create any new images
      @unique_image_paths = []
      @images.select { |image| image.is_new }.each do |new_image|
        next if @unique_image_paths.include? new_image.target # we only want to write each image once
        @unique_image_paths << new_image.target
        @file.output_stream.put_next_entry("word/#{new_image.target}")
        @file.output_stream.write ::File.read(new_image.path)
      end
    end

  end # ImageManager class
end # ODFReport module
