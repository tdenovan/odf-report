module ODFReport

  module Images

    IMAGE_DIR_NAME = "word/media"
    RELS_FILE = "word/_rels/document.xml.rels"

    def find_image_ids_and_paths(content)
      @images.each_pair do |image_name, path|

        if content.namespaces.include? 'xmlns' and content.xpath("//xmlns:Relationship").any? # Looking through word/_rels/document.xml.rels
          content.xpath("//xmlns:Relationship").each do |rel|
            @image_paths[rel.attr('Id')] = rel.attr('Target') # save image paths as a hash of 'id' => 'docx-image-path'
          end
        elsif content.namespaces.include? 'xmlns:w' and content.xpath("//w:drawing").any? # Looking through word/document.xml
          content.xpath("//w:drawing//wp:docPr[@title='#{image_name}']/following-sibling::*").xpath(".//a:blip", {'a' => "http://schemas.openxmlformats.org/drawingml/2006/main"}).each do |img|
            @image_ids[image_name] = img.attributes['embed'].value # save as a hash of 'image-name' => 'id'
          end
        end
      end # image looop
    end # find_and_replace_image_names method

    # Writes images to the zip file - this includes images that are not being replaced, as the File class was modified so that not files in the word/media directory are written (in case the files is being replaced in this method and would be overwritten)
    # CHANGED - previously, all files in the word/media directory were written to the zip. Then, in this method, any images being replaced were overwritten. That approach caused the zip to have an invalid header. The revised approach is to write all files in the word/media directory once, and substitute any images being replaced at the time of the first (and only) writing
    # TODO should also delete old images
    def create_images(file)
      
      # Enumerate all the files in the zip and write any that are in the media directory to the output buffer (which is used to generate the new zip file)
      file.read_files do |entry| # entry is a file entry in the zip
        if entry.name.include? IMAGE_DIR_NAME
          # Check if this is an image being replaced
          image_id = nil
          @image_paths.each { |id, docx_path| image_id = id if entry.name.include? docx_path }
          
          if @image_ids.key(image_id)
            replacement_path = @images[@image_ids.key(image_id)]
            data = ::File.read(replacement_path)
          else
            data = entry.get_input_stream { |is| data = is.sysread }            
          end
          
          file.output_stream.put_next_entry(entry.name)
          file.output_stream.write data
        end
      end
    end # create_images
    
    # Replaces image paths in the document.xml.rels file, so that the hrefs (paths) point to the new images we saved to the zip file
    # Note - the zip file corrupts if we simply replace the existing images, so we need to save the replacement images as new files, and update the paths
    def replace_image_paths(file, new_image_paths)
      file.update_file(RELS_FILE) do |txt|

        doc = Nokogiri::XML(txt)
        return txt unless doc.xpath("//xmlns:Relationship").any? 
        
        doc.xpath("//xmlns:Relationship").each do |rel|
          replaced_image = new_image_paths.select { |path| path[:old_path] == rel.attr('Target') }.first
          next if replaced_image.nil?
          rel.set_attribute('Target', replaced_image[:new_path])
        end

        txt.replace(doc.to_xml(:save_with => Nokogiri::XML::Node::SaveOptions::AS_XML))
      end
    end
    
    # Deletes old images
    def delete_old_images(file, zipfile_path)
      # Delete old images
      file.delete_files(*@new_image_paths.collect { |new_image| "word/" + new_image[:old_path]}, zipfile_path)
    end

    # newer versions of LibreOffice can't open files with duplicates image names
    def avoid_duplicate_image_names(content)

      nodes = content.xpath("//draw:frame[@draw:name]")

      nodes.each_with_index do |node, i|
        node.attribute('name').value = "pic_#{i}"
      end

    end

  end

end
