module ODFReport

  module Images

    IMAGE_DIR_NAME = "Pictures" # Should be word/media

    def find_image_name_matches(content)

      blip_id = blip_id || {} # Creating a hash of image names and linking them with their id
      @images.each_pair do |image_name, path|

        # Below we grab the image name. However, if a document doesn't have the image, it crashes
        blip_id[image_name] = content.xpath("//w:drawing//wp:docPr[@title=\"#{image_name}\"]/following-sibling::*").xpath("//a:blip", {'a' => "http://schemas.openxmlformats.org/drawingml/2006/main"}).attr('embed').value
        debugger

        if node = content.xpath("//draw:frame[@docPr:name='#{image_name}']/draw:image").first
          placeholder_path = node.attribute('href').value
          @image_names_replacements[path] = ::File.join(IMAGE_DIR_NAME, ::File.basename(placeholder_path))
        end
      end

    end

    def replace_images(file)

      return if @images.empty?

      @image_names_replacements.each_pair do |path, template_image|

        file.output_stream.put_next_entry(template_image)
        file.output_stream.write ::File.read(path)

      end

    end # replace_images

    # newer versions of LibreOffice can't open files with duplicates image names
    def avoid_duplicate_image_names(content)

      nodes = content.xpath("//draw:frame[@draw:name]")

      nodes.each_with_index do |node, i|
        node.attribute('name').value = "pic_#{i}"
      end

    end

  end

end
