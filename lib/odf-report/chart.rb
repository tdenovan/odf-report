module ODFReport

class Chart
  include Nested

  attr_accessor :name, :collection, :series, :title, :type, :legend, :labels, :file, :id, :target, :excel

  @@id_target = {}
  @@name_id = {}

  def initialize(opts)
    @name       = opts[:name]
    @collection = opts[:collection]

    @series     = opts[:series] || nil
    @type       = opts[:type]   || nil
    @colors     = opts[:colors] || nil

    @title      = opts[:title]  || :default
    @legend     = opts[:legend] || :default
    @labels     = opts[:labels] || :default

    @file       = opts[:file]   || nil
    @id         = ''
    @target     = ''
    @excel      = ''
  end

  def replace!(doc, filename, row = nil)

    if /_rels\/document/ === filename # Look into _rels/document.xml.rels

      doc.xpath("//xmlns:Relationship").each do |rel|
        @@id_target[rel.attr('Id')] = rel.attr('Target')
      end

    elsif /_rels\/chart/ === filename # Look into _rels/chart.xml.rels

      chart_name = @name
      chart_id = @@name_id[chart_name]
      chart_target = @@id_target[chart_id].split("\/").last

      if /#{Regexp.quote(chart_target)}/ === filename

        target = doc.xpath("//xmlns:Relationship").first['Target']
        current_path = "word" + target[2..-1]

        excel_file_data = @file.read_file(current_path)
        tmp_filename = "temp_#{current_path.split('/').last}"
        ::File.open(tmp_filename, "wb") {|f| f.write(excel_file_data) } # TODO fix the temporary file to actually reference a temporary file path in the tmp folder


        report = ODFReport::Report.new(tmp_filename) do |r|
           r.add_spreadsheet(@name, @collection, :series => @series, :title => @title, :type => @type)
        end

        replacement_path = current_path.split('/').last

        report.generate(replacement_path)

      end

    elsif /word\/document/ === filename # Look into word/document.xml

      @@name_id[@name] = doc.xpath("//w:drawing//wp:docPr[@title='#{@name}']/following-sibling::*").xpath(".//c:chart", {'c' => "http://schemas.openxmlformats.org/drawingml/2006/chart"}).attr('id').value

    elsif /charts\/chart/ === filename # Look through chart.xml

      if filename.include? @@id_target[@@name_id[@name]]

        determine_type(doc)

        case @type

        when 'pie', 'doughnut' # For Pie/Doughnut Charts

          @series = [@series] unless @series.is_a? Array
          @collection.each { |k, v| @collection[k] = [v] } unless @collection.values.first.is_a? Array

          add_series(doc, @type)

          add_color(doc)

          add_data(doc)

          add_options(doc)

        when 'waterfall' # For Waterfall charts

          etl_waterfall(doc) unless @series == ['Fill', 'Base', 'Rise+', 'Rise-', 'Fall+', 'Fall-']

          add_series(doc)

          add_color(doc)

          add_data(doc)

          add_options(doc)

        when 'bar', 'column' # For Bar/Column Charts

          if @series.length < @collection.values.first.length
            @series << rand(65..91).chr until @series.length == @collection.values.first.length
          elsif @series.length < @collection.values.first.length
            @series.pop until @series.length == @collection.values.first.length
          end

          add_series(doc)

          add_color(doc)

          add_data(doc)

          add_options(doc)

        when 'line' # For Line Charts

          if @series.length < @collection.values.first.length
            @series << rand(65..91).chr until @series.length == @collection.values.first.length
          elsif @series.length < @collection.values.first.length
            @series.pop until @series.length == @collection.values.first.length
          end

          add_series(doc, @type)

          add_color(doc)

          add_data(doc)

          add_options(doc)

        end

      end

    end

  end

  private

  def determine_type(doc) # Still a prototype since sometimes the spreadsheet is done first

    return unless @type.nil?
    @type = 'bar' if doc.xpath("//c:barChart").any?
    @type = 'pie' if doc.xpath("//c:pieChart").any?
    @type = 'doughnut' if doc.xpath("//c:doughnutChart").any?
    @type = 'waterfall' if doc.xpath("//c:grouping").first['val'] == 'clustered'

  end

  def add_series(doc, type = 'bar')

    rows = @collection.length

    doc.xpath("//c:ser").remove

    @series.length.times do |i|
      series_temp = "<c:ser><c:idx val=\"#{i}\"/><c:order val=\"#{i}\"/><c:tx><c:strRef><c:f>Sheet1!$#{(66 + i).chr}$1</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>New Series</c:v></c:pt></c:strCache></c:strRef></c:tx><c:invertIfNegative val=\"0\"/><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$#{rows + 1}</c:f><c:strCache><c:ptCount val=\"1\"/></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$#{(66 + i).chr}$2:$#{(66 + i).chr}$#{rows + 1}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"1\"/></c:numCache></c:numRef></c:val></c:ser>"
      doc.xpath("//c:#{type}Chart").first.add_child(series_temp)
    end

    length = @collection.length

    doc.xpath("//c:ser").each do |series|
      column_idx = 0

      until series.xpath(".//c:cat//c:pt").length == length
        column_temp = "<c:pt idx=\"#{column_idx}\"><c:v>New Column</c:v></c:pt>"
        value_temp = "<c:pt idx=\"#{column_idx}\"><c:v>0.0</c:v></c:pt>"

        series.xpath(".//c:strCache").last.add_child(column_temp)
        series.xpath(".//c:numCache").last.add_child(value_temp)

        column_idx += 1
      end
    end

    doc.xpath("//c:ser//c:v").first.content = @series if @series.class == String
    doc.xpath("//c:tx//c:v").each_with_index { |name, index| name.content = @series[index] } if @series.class == Array


  end

  def add_data(doc)

    doc.xpath("//c:cat//c:v").each_with_index do |node, index|

      until index < @collection.length
        index -= @collection.length
      end

      node.content = @collection.keys[index]

    end

    doc.xpath("//c:val//c:v").each_with_index do |node, index|
      series = 0

      until index < @collection.length
        index -= @collection.length
        series += 1
      end

      node.content = @collection.values[index][series]

    end

  end

  def add_color(doc)

    return if @colors.nil?

    @color_fill = []

    @colors.each do |color|

      if color.nil? or color.to_i > 6

        @color_fill << nil

      elsif color.to_i.zero?

        fill = "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>"
        @color_fill << fill

      elsif color.is_a? Fixnum

        fill = "<c:spPr><a:solidFill><a:schemeClr val=\"accent#{color}\"/></a:solidFill></c:spPr>"
        fill = "<c:spPr><a:ln><a:solidFill><a:schemeClr val=\"accent#{color}\"/></a:solidFill></a:ln></c:spPr><c:marker><c:spPr><a:solidFill><a:schemeClr val=\"accent#{color}\"></a:schemeClr></a:solidFill><a:ln><a:solidFill><a:schemeClr val=\"accent1\"></a:schemeClr></a:solidFill></a:ln></c:spPr></c:marker>" if @type == 'line'
        @color_fill << fill

      elsif color.is_a? Float

        lightness  = [nil, {mod: 20000, off: 80000}, {mod: 40000, off: 60000}, {mod: 60000, off: 40000}, {mod: 75000, off: nil}, {mod: 50000, off: nil}]
        lum_index  = color.to_s[-1].to_i

        if lum_index.zero?
          fill = "<c:spPr><a:solidFill><a:schemeClr val=\"accent#{color.to_i}\"/></a:solidFill></c:spPr>"
          fill = "<c:spPr><a:ln><a:solidFill><a:schemeClr val=\"accent#{color}\"/></a:solidFill></a:ln></c:spPr><c:marker><c:spPr><a:solidFill><a:schemeClr val=\"accent#{color}\"></a:schemeClr></a:solidFill><a:ln><a:solidFill><a:schemeClr val=\"accent1\"></a:schemeClr></a:solidFill></a:ln></c:spPr></c:marker>" if @type == 'line'
          @color_fill << fill
          next
        elsif lum_index > 5
          @color_fill << nil
          next
        end

        fill = "<c:spPr><a:solidFill><a:schemeClr val=\"accent#{color.to_i}\"><a:lumMod val=\"#{lightness[lum_index][:mod]}\"/><a:lumOff val=\"#{lightness[lum_index][:off]}\"/></a:schemeClr></a:solidFill></c:spPr>" if lum_index <= 3
        fill = "<c:spPr><a:solidFill><a:schemeClr val=\"accent#{color.to_i}\"><a:lumMod val=\"#{lightness[lum_index][:mod]}\"/></a:schemeClr></a:solidFill></c:spPr>" if lum_index > 3
        fill = "<c:spPr><a:ln><a:solidFill><a:schemeClr val=\"accent#{color.to_i}\"><a:lumMod val=\"#{lightness[lum_index][:mod]}\"/><a:lumOff val=\"#{lightness[lum_index][:off]}\"/></a:schemeClr></a:solidFill></a:ln></c:spPr><c:marker><c:spPr><a:solidFill><a:schemeClr val=\"accent#{color.to_i}\"><a:lumMod val=\"#{lightness[lum_index][:mod]}\"/><a:lumOff val=\"#{lightness[lum_index][:off]}\"/></a:schemeClr></a:solidFill><a:ln><a:solidFill><a:schemeClr val=\"accent#{color.to_i}\"><a:lumMod val=\"#{lightness[lum_index][:mod]}\"/><a:lumOff val=\"#{lightness[lum_index][:off]}\"/></a:schemeClr></a:solidFill></a:ln></c:spPr></c:marker>" if lum_index <= 3 and @type == 'line'
        fill = "<c:spPr><a:ln><a:solidFill><a:schemeClr val=\"accent#{color.to_i}\"><a:lumMod val=\"#{lightness[lum_index][:mod]}\"/></a:schemeClr></a:solidFill></a:ln></c:spPr><c:marker><c:spPr><a:solidFill><a:schemeClr val=\"accent#{color.to_i}\"><a:lumMod val=\"#{lightness[lum_index][:mod]}\"/></a:schemeClr></a:solidFill><a:ln><a:solidFill><a:schemeClr val=\"accent#{color.to_i}\"><a:lumMod val=\"#{lightness[lum_index][:mod]}\"/></a:schemeClr></a:solidFill></a:ln></c:spPr></c:marker>" if lum_index > 3 and @type == 'line'
        @color_fill << fill

      end

    end

    case @type
    when 'pie', 'doughnut'

      @color_fill.each_with_index do |color, index|
        next if color.nil?
        fill = "<c:dPt><c:idx val=\"#{index}\"/><c:bubble3D val=\"0\"/>#{color}</c:dPt>"
        doc.xpath("//c:ser").first.add_child(fill)
      end

    else

      doc.xpath("//c:ser").each_with_index do |series, index|
        next if @color_fill[index].nil?
        series.xpath(".//c:v").first.content = @series[index]
        series.add_child(@color_fill[index])
      end

    end

  end

  def add_options(doc)

    doc.xpath("//c:autoTitleDeleted").first['val'] = 1 unless @title == :default
    doc.xpath("//c:title").remove unless @title == :default
    doc.xpath("//c:dLbls").remove unless @labels == :default
    doc.xpath("//c:legend").remove unless @legend == :default

    if @title.is_a? String
      title_temp = '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>New Title</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title>'
      doc.xpath("//c:chart").first.add_child(title_temp)
      doc.xpath("//a:t").first.content = @title
    end

    if @labels == :enabled and (@type == 'pie' or @type == 'doughnut')
      labels_temp = "<c:dLbls><c:dLblPos val=\"outEnd\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/><c:showCatName val=\"1\"/><c:showSerName val=\"0\"/><c:showPercent val=\"1\"/><c:showBubbleSize val=\"0\"/><c:separator/><c:showLeaderLines val=\"1\"/></c:dLbls>"
      doc.xpath("//c:ser").first.add_child(labels_temp)
      doc.xpath("//c:dLblPos").first.remove if @type == 'doughnut'
    end

    if @legend == :enabled
      legend_temp = "<c:legend><c:legendPos val=\"b\"/><c:layout/><c:overlay val=\"0\"/></c:legend>"
      doc.xpath("//c:chart").first.add_child(legend_temp)
    end

    case @type
    when 'bar', 'column', 'waterfall'
      doc.xpath("//c:gapWidth").first['val'] = 50
      doc.xpath("//c:barChart").first.add_child("<c:overlap val=\"-50\"/>")
      doc.xpath("//c:catAx//c:tickLblPos").first['val'] = "low" unless doc.xpath("//c:catAx//c:delete").first['val'] == '1' or doc.xpath("//c:valAx//c:delete").first['val'] == '1'
    end

  end

  def etl_waterfall(doc)
    input = @collection.values
    output = {
      'fill' => [],
      'base' => [],
      'pro+' => [],
      'pro-' => [],
      'con+' => [],
      'con-' => []
    }

    @series = ['Fill', 'Base', 'Rise+', 'Rise-', 'Fall+', 'Fall-']

    @colors = [0, 1, 5.3, 5.3, 3.3, 3.3]

    doc.xpath('//c:legend').remove

    sum = 0

    return if @collection.values.first.is_a? Array

    input.each_with_index do |num, index|

      if index == 0
        output['base'] << num
      elsif index == input.length - 1
        output['base'] << sum
      elsif num >= 0 and sum >= 0
        output['fill'] << sum
        output['pro+'] << num
      elsif num >= 0 and sum < 0 and sum + num < 0
        output['fill'] << sum + num
        output['pro-'] << - num
      elsif num >= 0 and sum < 0 and sum + num >= 0
        output['pro+'] << sum + num
        output['pro-'] << sum
      elsif num < 0 and sum < 0
        output['fill'] << sum
        output['con-'] << num
      elsif num < 0 and sum >= 0 and sum + num >= 0
        output['fill'] << sum + num
        output['con+'] << - num
      elsif num < 0 and sum >= 0 and sum + num < 0
        output['con+'] << sum
        output['con-'] << sum + num
      end

      length = output.values.sort_by { |x| x.length }.last.length
      output.values.each { |arr| arr << 0 if arr.length < length }
      sum += num
    end

    output.values.transpose.each_with_index { |array, index| @collection[@collection.keys[index]] = array }

  end

end

end
