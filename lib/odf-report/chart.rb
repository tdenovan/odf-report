module ODFReport

class Chart
  include Nested

  @@id_target = {}
  @@name_id = {}

  def initialize(opts)
    @name             = opts[:name]
    @collection       = opts[:collection]

    @series           = opts[:series] || nil
    @title            = opts[:title] || nil

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

        excel_file = doc.xpath("//xmlns:Relationship").first['Target']

        ##########

        # Make a function here to go into the excel file and edit the xml files in there

        ##########

      end

    elsif /word\/document/ === filename # Look into word/document.xml

      @@name_id[@name] = doc.xpath("//w:drawing//wp:docPr[@title='#{@name}']/following-sibling::*").xpath(".//c:chart", {'c' => "http://schemas.openxmlformats.org/drawingml/2006/chart"}).attr('id').value

    elsif /charts\/chart/ === filename # Look through chart.xml

      if filename.include? @@id_target[@@name_id[@name]]

        if doc.namespaces.include? 'xmlns:c' and doc.xpath("//c:pieChart").any? or doc.xpath("//c:doughnutChart").any? # For Pie/Doughnut Charts

          @series = [@title] unless @series.is_a? Array
          @collection.each { |k, v| @collection[k] = [v] } unless @collection.values.first.is_a? Array

          type = 'pie' if doc.xpath("//c:pieChart").any?
          type = 'doughnut' if doc.xpath("//c:doughnutChart").any?

          add_series(doc, type)

          add_data(doc)

        elsif doc.xpath("//c:grouping").attr('val').value == 'stacked' # For Waterfall charts

          etl_waterfall

          add_series(doc)

          doc.xpath("//c:ser").each_with_index do |series, index|
            series.xpath(".//c:v").first.content = @series[index]
            series.add_child(@color_fill[index])
          end

          add_data(doc)

        elsif doc.namespaces.include? 'xmlns:c' and doc.xpath("//c:barChart").any? # For Bar/Column Charts

          if @series.length < @collection.values.first.length
            @series << rand(65..91).chr until @series.length == @collection.values.first.length
          elsif @series.length < @collection.values.first.length
            @series.pop until @series.length == @collection.values.first.length
          end

          add_series(doc)

          add_data(doc)

        end

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

    doc.xpath("//c:title").remove
    if !!@title
      title_temp = '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>New Title</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title>'
      doc.xpath("//c:chart").first.add_child(title_temp)
      doc.xpath("//a:t").first.content = @title
    end

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

  def etl_waterfall
    input = @collection.values
    output = {
      'fill' => [],
      'base' => [],
      'pro+' => [],
      'pro-' => [],
      'con+' => [],
      'con-' => []
    }

    @series = ['Fill', 'Base', 'Rise', 'Rise', 'Fall', 'Fall']

    @color_fill = [
      "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>",
      "<c:spPr><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill></c:spPr>",
      "<c:spPr><a:solidFill><a:schemeClr val=\"accent6\"/></a:solidFill></c:spPr>",
      "<c:spPr><a:solidFill><a:schemeClr val=\"accent6\"/></a:solidFill></c:spPr>",
      "<c:spPr><a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill></c:spPr>",
      "<c:spPr><a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill></c:spPr>"
    ]

    sum = 0

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
