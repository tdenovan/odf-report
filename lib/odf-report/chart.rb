module ODFReport

class Chart
  include Nested

  attr_accessor :name, :collection, :series, :title, :type, :legend, :labels, :file, :id, :target, :excel

  @@chart_ids = {}

  def initialize(opts)
    @name       = opts[:name]
    @collection = opts[:collection]

    @series     = opts[:series] || nil
    @type       = opts[:type]   || nil
    @colors     = opts[:colors] || nil

    @title      = opts[:title]  || :default
    @legend     = opts[:legend] || :default
    @labels     = opts[:labels] || :default
    @x_axis     = opts[:x_axis] || :default
    @y_axis     = opts[:y_axis] || :default
    @f_size     = opts[:font_size] || :default
    @g_line     = opts[:grid_lines] || :default

    @file       = opts[:file]   || nil
    @id         = ''
    @target     = ''
    @excel      = ''
  end

  def replace!(doc, filename, row = nil)

    if /_rels\/document/ === filename # Look into _rels/document.xml.rels

      doc.xpath("//xmlns:Relationship").each do |rel|
        id = rel.attr('Id')
        target = rel.attr('Target')

        @@chart_ids[id] = {} if @@chart_ids[target].nil?
        @@chart_ids[id].merge!({:target => target})
      end

    elsif /word\/document/ === filename # Look into word/document.xml
      ids = doc.xpath("//w:drawing//wp:docPr[@title='#{@name}']/following-sibling::*").xpath(".//c:chart", {'c' => "http://schemas.openxmlformats.org/drawingml/2006/chart"})
      return if ids.empty?

      ids.map { |id| id.attributes['id'].value }.each do |id|
        @@chart_ids[id] = {} if @@chart_ids[id].nil?
        @@chart_ids[id].merge!({:name => @name})
      end

      doc.xpath("//w:drawing//wp:docPr[@title='#{@name}']/ancestor::*[w:drawing]").remove if @collection.values.flatten.map {|v| v.abs}.sum.zero?

    elsif /_rels\/chart/ === filename # Look into _rels/chart.xml.rels

      return if @@chart_ids.values.select { |v| v[:name] == @name and /#{Regexp.quote(v[:target].split("\/").last)}/ === filename }.empty?

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

    elsif /charts\/chart/ === filename # Look through chart.xml

      return if @@chart_ids.values.select { |v| v[:name] == @name and /#{Regexp.quote(v[:target])}/ === filename }.empty?

      case @type

      when 'pie', 'doughnut' # For Pie/Doughnut Charts

        @series = [@series] unless @series.is_a? Array
        @collection.each { |k, v| @collection[k] = [v] } unless @collection.values.first.is_a? Array

        add_series(doc, @type)

        add_color(doc)

        add_data(doc)

        add_options(doc)

      when 'waterfall' # For Waterfall charts

        etl_waterfall(doc)

        add_series(doc)

        add_color(doc)

        add_data(doc)

        add_options(doc)

      when 'bar', 'column' # For Bar/Column Charts

        if @series.length < @collection.values.first.length
          @series << rand(65..91).chr until @series.length == @collection.values.first.length
        elsif @series.length > @collection.values.first.length
          @series.pop until @series.length == @collection.values.first.length
        end

        add_series(doc)

        add_color(doc)

        add_data(doc)

        add_options(doc)

      when 'line' # For Line Charts

        if @series.length < @collection.values.first.length
          @series << rand(65..91).chr until @series.length == @collection.values.first.length
        elsif @series.length > @collection.values.first.length
          @series.pop until @series.length == @collection.values.first.length
        end

        add_series(doc, @type)

        add_color(doc)

        add_data(doc)

        add_options(doc)

      end

    end

  end

  private

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

      if color.nil?

        color_temp = nil

      elsif color.is_a? String and color =~ /#\h{3,6}/

        color_temp = "<a:solidFill><a:srgbClr val=\"#{color[1..6]}\"/></a:solidFill>"

      elsif color.to_i.zero?

        color_temp = "<a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/>"

      elsif color.is_a? Fixnum

        color_temp = "<a:solidFill><a:schemeClr val=\"accent#{color}\"/></a:solidFill>"

      elsif color.is_a? Float

        lightness  = [nil, {mod: 20000, off: 80000}, {mod: 40000, off: 60000}, {mod: 60000, off: 40000}, {mod: 75000, off: nil}, {mod: 50000, off: nil}]
        lum_index  = color.to_s[-1].to_i

        case lum_index
        when 0 then color_temp = "<a:solidFill><a:schemeClr val=\"accent#{color.to_i}\"/></a:solidFill>"
        when 1, 2, 3 then color_temp = "<a:solidFill><a:schemeClr val=\"accent#{color.to_i}\"><a:lumMod val=\"#{lightness[lum_index][:mod]}\"/><a:lumOff val=\"#{lightness[lum_index][:off]}\"/></a:schemeClr></a:solidFill>"
        when 4, 5 then color_temp = "<a:solidFill><a:schemeClr val=\"accent#{color.to_i}\"><a:lumMod val=\"#{lightness[lum_index][:mod]}\"/></a:schemeClr></a:solidFill>"
        else
          color_temp = nil
        end

      else

        color_temp = nil

      end


      if color_temp.nil?
      elsif @type == 'line'
        color_temp = "<c:spPr><a:ln w=\"12700\">#{color_temp}</a:ln></c:spPr><c:marker><c:spPr>#{color_temp}<a:ln w=\"12700\">#{color_temp}</a:ln></c:spPr></c:marker>"
      else
        color_temp = "<c:spPr>#{color_temp}</c:spPr>"
      end

      @color_fill << color_temp

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

    doc.xpath("//c:autoTitleDeleted").first['val'] = 1
    doc.xpath("//c:title").remove unless @title == :default
    doc.xpath("//c:dLbls").remove unless @labels == :default
    doc.xpath("//c:legend").remove unless @legend == :default
    doc.xpath("//c:chartSpace/c:spPr").remove

    doc.xpath("//c:chartSpace").first.add_child("<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>")

    if @title.is_a? String
      title_temp = '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>New Title</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title>'
      doc.xpath("//c:chart").first.add_child(title_temp)
      doc.xpath("//a:t").first.content = @title
    end

    if @labels == :enabled and (@type == 'pie' or @type == 'doughnut')
      labels_temp = "<c:dLbls><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"800\"/></a:pPr><a:endParaRPr lang=\"en-US\"/></a:p></c:txPr><c:dLblPos val=\"outEnd\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/><c:showCatName val=\"1\"/><c:showSerName val=\"0\"/><c:showPercent val=\"1\"/><c:showBubbleSize val=\"0\"/><c:separator/><c:showLeaderLines val=\"1\"/></c:dLbls>"
      doc.xpath("//c:ser").first.add_child(labels_temp)
      doc.xpath("//c:dLblPos").first.remove if @type == 'doughnut'
    end

    if @legend == :enabled
      legend_temp = "<c:legend><c:legendPos val=\"b\"/><c:layout/><c:overlay val=\"0\"/></c:legend>"
      doc.xpath("//c:chart").first.add_child(legend_temp)
    end

    unless @f_size == :default
      @f_size = 8 if @f_size == :enabled
      font_temp = "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"#{@f_size}00\"><a:latin typeface=\"Georgia\"/><a:cs typeface=\"Georgia\"/></a:defRPr></a:pPr><a:endParaRPr lang=\"en-US\"/></a:p></c:txPr>"
      doc.xpath("//c:txPr").remove
      doc.xpath("//c:chart").first.add_next_sibling(font_temp)
    end


    case @type
    when 'pie', 'doughnut'
      line_temp = "<a:ln><a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill></a:ln>"
      doc.xpath("//c:ser//c:tx").first.add_next_sibling("<c:spPr></c:spPr>")
      doc.xpath("//c:dPt//c:spPr").each { |xml| xml.add_child(line_temp) }

    when 'bar', 'column', 'waterfall', 'line'

      unless @g_line == :default
        grid_temp = "<c:majorGridlines><c:spPr><a:ln cmpd=\"sng\" w=\"3175\"><a:solidFill><a:schemeClr val=\"bg1\"><a:lumMod val=\"75000\"/></a:schemeClr></a:solidFill></a:ln></c:spPr></c:majorGridlines>"
        doc.xpath("//c:valAx//c:majorGridlines").remove
        doc.xpath("//c:valAx//c:axPos").first.add_next_sibling(grid_temp) unless @g_line == :disabled
      end

      unless @x_axis == :default

        doc.xpath("//c:catAx//c:delete").first['val'] = 0 if @x_axis == :enabled
        doc.xpath("//c:catAx//c:majorTickMark").first['val'] = 'none' if @x_axis == :disabled
        doc.xpath("//c:catAx//c:tickLblPos").first['val'] = 'none' if @x_axis == :disabled


        max_length = @collection.keys.sort_by { |cat| cat.length }.last.length
        if @x_axis == :enabled and max_length > 25
          font_temp = "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"600\"/></a:pPr><a:endParaRPr lang=\"en-US\"/></a:p></c:txPr>"
          doc.xpath("//c:tickLblPos").first.add_next_sibling(font_temp)
        end

      end

      unless @y_axis == :default

        doc.xpath("//c:valAx//c:delete").first['val'] = 0 if @y_axis == :enabled
        doc.xpath("//c:valAx//c:delete").first['val'] = 1 if @y_axis == :disabled

        case @collection.values.flatten.max.to_i.to_s.length
        when 4, 5, 6 then scale = 'thousands'
        when 7, 8, 9 then scale = 'millions'
        when 10, 11, 12 then scale = 'billions'
        end

        unless scale.nil?
          scale_temp = "<c:dispUnits><c:builtInUnit val=\"#{scale}\"/><c:dispUnitsLbl><c:layout/></c:dispUnitsLbl></c:dispUnits>"
          doc.xpath("//c:valAx//c:dispUnits").remove
          doc.xpath("//c:valAx").first.add_child(scale_temp)
          doc.xpath("//c:valAx//c:dispUnitsLbl").remove if @y_axis == :noscale
        end

        max = @collection.values.flatten.max
        min = @collection.values.flatten.min

        max = @collection.values.map {|x| x.inject(:+) }.max if @type == 'waterfall'
        min = @collection.values.map {|x| x.inject(:+) }.min if @type == 'waterfall'

        def roundup(x)
          x >= 1 ? y = 10 ** Math.log10(x).to_i : y = 10 ** (Math.log10(x).to_i - 1)
          return x if x % y == 0   # already a factor of 10
          return x + y - (x % y)  # go to nearest factor 10
        end

        max <= 0 ? max = 0 : max = roundup(max)
        min >= 0 ? min = 0 : min = -roundup(-min)

        if max > 0 and min < 0

          max_10 = Math.log10(max).to_i
          min_10 = Math.log10(-min).to_i

          if max_10 > min_10
            min = -(10 ** max_10)
          elsif max_10 < min_10
            max = 10 ** min_10
          end

        end

        max = max * 3 / 4 if max.to_s[0] == '2' and @collection.values.flatten.max.to_s[1].to_i < 5

        tick = (max - min) / 2

        max_min_temp = "<c:max val=\"#{max}\"/><c:min val=\"#{min}\"/>"
        tick_temp = "<c:majorUnit val=\"#{tick}\"/>"

        doc.xpath("//c:valAx//c:max").remove
        doc.xpath("//c:valAx//c:min").remove
        doc.xpath("//c:valAx//c:majorUnit").remove

        doc.xpath("//c:valAx//c:scaling").first.add_child(max_min_temp)
        doc.xpath("//c:valAx").first.add_child(tick_temp)

      end


      doc.xpath("//c:catAx//c:tickLblPos").first['val'] = "low" unless doc.xpath("//c:catAx//c:delete").first['val'] == '1' or doc.xpath("//c:valAx//c:delete").first['val'] == '1'

      return if @type == 'line'
      doc.xpath("//c:gapWidth").first['val'] = 50

      return if @type == 'waterfall'
      doc.xpath("//c:barChart//c:overlap").remove
      doc.xpath("//c:barChart").first.add_child("<c:overlap val=\"-50\"/>")

    end

  end

  def etl_waterfall(doc)

    return if @series == ['Fill', 'Base', 'Rise+', 'Rise-', 'Fall+', 'Fall-']

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

    @colors = [0, 1, '#FFB600', '#FFB600', 3.3, 3.3]

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
