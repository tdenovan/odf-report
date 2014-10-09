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

        file.each do |entry|

          # next if entry.directory?

          entry.get_input_stream do |is|
            data = is.sysread

            if content_files.include?(entry.name) or content_files.select { |filename| filename.is_a? Regexp and filename =~ entry.name }.count >= 1
              yield data, entry.name
            end

            unless entry.name.include? Images::IMAGE_DIR_NAME # do not write images, these will be written later
              @output_stream.put_next_entry(entry.name)
              @output_stream.write data
            end

          end

        end

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

    def delete_files(*paths, zipfile_path)
      Zip::File.open(zipfile_path, Zip::File::CREATE) do |file|
        debugger
        paths.each { |path| file.remove(path) }
      end
    end

    def data
      @buffer.string
    end

  end
end
