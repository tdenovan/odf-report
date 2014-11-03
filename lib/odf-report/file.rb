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
      excluded_paths = [ImageManager::IMAGE_DIR_NAME, RelationshipManager::RELATIONSHIP_FILE, ChartManager::CHART_DIR_NAME]

      Zip::File.open(@template) do |file|

        file.each do |entry|

          # next if entry.directory?

          entry.get_input_stream do |is|
            data = is.sysread

            if content_files.include?(entry.name) or content_files.select { |filename| filename.is_a? Regexp and filename =~ entry.name }.count >= 1
              yield data, entry.name
            end

            # Check if this is an excluded path - excluded paths (e.g. the Relationship file) are written separately
            unless excluded_paths.select { |path| entry.name.include? path }.count >= 1
              @output_stream.put_next_entry(entry.name)
              @output_stream.write data
            end

          end

        end

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
      Zip::File.open(@template) do |file|
        file.each do |entry|
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
