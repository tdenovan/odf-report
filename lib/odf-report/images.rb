module ODFReport

  module Images

    IMAGE_DIR_NAME = "word/media" # Should be word/media

    def find_image_name_matches(content)


      @images.each_pair do |image_name, path|

        if content.namespaces.include? 'xmlns' and content.xpath("//xmlns:Relationship").any? # Looking through word/_rels/document.xml.rels

          content.xpath("//xmlns:Relationship").each do |rel|
            @image_id_paths[rel.attr('Id')] = rel.attr('Target')
          end

        elsif content.namespaces.include? 'xmlns:w' and content.xpath("//w:drawing").any? # Looking through word/document.xml

          content.xpath("//w:drawing//wp:docPr[@title='#{image_name}']/following-sibling::*").xpath(".//a:blip", {'a' => "http://schemas.openxmlformats.org/drawingml/2006/main"}).each do |img|
            @image_name_id[image_name] = img.attributes['embed'].value
          end

          # @image_name_id[image_name] = content.xpath("//w:drawing//wp:docPr[@title='#{image_name}']/following-sibling::*").xpath(".//a:blip", {'a' => "http://schemas.openxmlformats.org/drawingml/2006/main"}).attributes['embed'].value

        end



        # if content.xpath("//Relationships").any?
        #   debugger
        #   if node = content.xpath("//draw:frame[@docPr:name='#{image_name}']/draw:image").first
        #     placeholder_path = node.attribute('href').value
        #     @image_names_replacements[path] = ::File.join(IMAGE_DIR_NAME, ::File.basename(placeholder_path))
        #   end
        # end
      end

    end

    def replace_images(file)

      # return if @images.empty?

      @images.each do |name, path|
        @image_names_replacements[path] = ::File.join(IMAGE_DIR_NAME, ::File.basename(@image_id_paths[@image_name_id[name]]))
      end

      @image_names_replacements.each do |path, template_image|
      # ************************
        file.output_stream.put_next_entry(template_image)
        file.output_stream.write ::File.read(path)
        # Code not working how we want it. Must replace method
      # ************************
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
