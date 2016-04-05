module ODFReport
  class File

    attr_accessor :output_stream

    def initialize(template)
      raise "Template [#{template}] not found." unless ::File.exists? template
      @template = template
    end

    def update_content
      @buffer = Zip::OutputStream.write_buffer do |out|
        @output_stream = out
        yield self
      end

    end

    def update_files(*content_files, &block)

      Zip::File.open(@template) do |file|
        parsed_files = []
        content_files.each do |content_file|
          file.each do |entry|
            file_matches_searched_file = content_file === entry.name
            file_matches_searched_file = content_file =~ entry.name if file_matches_searched_file == false and content_file.is_a? Regexp
            if file_matches_searched_file and not parsed_files.include?(entry.name)
              parsed_files.push(entry.name)
              entry.get_input_stream do |is|
                data = is.sysread
                yield data, entry.name

                # Check if this is an excluded path - excluded paths (e.g. the Relationship file) are written separately
                update_files_output_stream entry, data
              end
            end
          end
        end

        file.each do |entry|
          unless parsed_files.include?(entry.name)
            entry.get_input_stream do |is|
              data = is.sysread
              # Check if this is an excluded path - excluded paths (e.g. the Relationship file) are written separately
              update_files_output_stream entry, data
            end

          end

        end

      end

    end

    def update_files_output_stream(entry, data)
      excluded_paths = [ImageManager::IMAGE_DIR_NAME, RelationshipManager::RELATIONSHIP_FILES[:ppt], RelationshipManager::RELATIONSHIP_FILES[:doc], ChartManager::CHART_DIR_NAME]
      unless excluded_paths.select { |path| entry.name.include? path }.count >= 1
        @output_stream.put_next_entry(entry.name)
        @output_stream.write data
      end
    end

    # Updates a specific file in the zip
    def update_file(filename, &block)
      Zip::File.open(@template) do |file|
        data = file.get_entry(filename).get_input_stream.sysread
        doc = Nokogiri::XML(data)
        yield doc
        data.replace(doc.to_xml(:save_with => Nokogiri::XML::Node::SaveOptions::AS_XML))

        file.remove(filename)
        os = file.get_output_stream(filename)
        os.write(data)
        os.close
      end
    end


    # Returns access to a file stream
    def read_files(&block)
      Zip::File.open(@template) do |files|
        files.each do |entry|
          yield entry
        end
      end
    end # update file stream

    # Reads a specific file - returns the stream
    def read_file(filename)
      file_data = nil
      Zip::File.open(@template) do |file|
        file_data = file.get_entry(filename).get_input_stream.sysread
      end

      file_data
    end

    def delete_files(*paths, zipfile_path)
      Zip::File.open(zipfile_path, Zip::File::CREATE) do |file|
        paths.each { |path| file.remove(path) }
      end
    end

    def data
      @buffer.string
    end

  end
end
