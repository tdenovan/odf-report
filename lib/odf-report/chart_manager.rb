module ODFReport

  # This class is responsible for managing the list of images in the document that will be replaced
  # It also keeps a record of new images (e.g. repeated images in tables) that need to be written
  # This class is effectively a singleton
  class ChartManager
    attr_accessor :existing_charts, :altered_charts

    CHART_DIR_NAME = "word/embedding"

    def initialize(relationship_manager, file_manager)
      @relationship_manager = relationship_manager
      @charts = []
      @file = file_manager
    end

    def add_charts(chart_name, collection, opts={})
      opts.merge!(:name => chart_name, :collection => collection, file: @file)
      chart = Chart.new(opts)
      @charts << chart
    end

    def find_chart_ids(doc, filename)

      if filename == 'word/document.xml'
        @charts.each do |chart|
          chart.id = doc.xpath("//w:drawing//wp:docPr[@title='#{chart.name}']/following-sibling::*").xpath(".//c:chart", {'c' => "http://schemas.openxmlformats.org/drawingml/2006/chart"}).attr('id').value
        end
      end

      return unless filename.include? 'charts/_rels'

      @charts.each do |chart|
        target = @relationship_manager.get_relationship(chart.id)[:target].split("/").last
        next unless /#{Regexp.quote(target)}/ === filename
        excel_path = doc.xpath("//xmlns:Relationship").first['Target']
        chart.excel = 'word' + excel_path[2..-1]
      end

    end

    def write_charts

      @file.read_files do |entry| # entry is a file entry in the zip

        next unless entry.name.include? CHART_DIR_NAME
        current_chart = @charts.select { |chart| entry.name == chart.excel }.first

        if current_chart.nil?

          data = ''
          entry.get_input_stream { |is| data = is.sysread }
          @file.output_stream.put_next_entry(entry.name)
          @file.output_stream.write data

        elsif

          replacement_path = current_chart.excel.split('/').last

          data = ::File.read(replacement_path)
          @file.output_stream.put_next_entry(entry.name)
          @file.output_stream.write data
          ::File.delete replacement_path
          ::File.delete "temp_#{replacement_path.split('/').last}"

        end

      end

    end

  end # ImageManager class
end # ODFReport module