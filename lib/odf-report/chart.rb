module ODFReport

class Chart
  include Nested

  @@id_target = {}
  @@name_id = {}

  def initialize(opts)
    @name             = opts[:name]
    @collection       = opts[:collection]

    @series_name      = opts[:series_name] || ''
    @chart_type       = opts[:chart_type] || ''

  end

  def replace!(doc, filename, row = nil)

    if doc.namespaces.include? 'xmlns' and doc.xpath("//xmlns:Relationship").any? # Look into _rels/document.xml.rels

      doc.xpath("//xmlns:Relationship").each do |rel|
        @@id_target[rel.attr('Id')] = rel.attr('Target')
      end

    elsif doc.namespaces.include? 'xmlns:w' and doc.xpath("//w:drawing").any? # Look into word/document.xml

      @@name_id[@name] = doc.xpath("//w:drawing//wp:docPr[@title='#{@name}']/following-sibling::*").xpath(".//c:chart", {'c' => "http://schemas.openxmlformats.org/drawingml/2006/chart"}).attr('id').value

    elsif doc.namespaces.include? 'xmlns:c' # Look through chart.xml

      if filename.include? @@id_target[@@name_id[@name]]

        if doc.namespaces.include? 'xmlns:c' and doc.xpath("//c:pieChart").any? # For Pie Charts

          length = @collection.length
          doc.xpath("//c:cat").xpath(".//c:pt").remove
          doc.xpath("//c:val").xpath(".//c:pt").remove
          column_idx = 0

          until doc.xpath(".//c:cat").xpath(".//c:pt").length == length
            column_temp = "<c:pt idx=\"#{column_idx}\"><c:v>New Column</c:v></c:pt>"
            value_temp = "<c:pt idx=\"#{column_idx}\"><c:v>0.0</c:v></c:pt>"

            doc.xpath(".//c:strCache").last.add_child(column_temp)
            doc.xpath(".//c:numCache").last.add_child(value_temp)

            column_idx += 1
          end

          doc.xpath("//c:cat//c:v").each_with_index do |node, index|
            node.content = @collection.keys[index]
          end

          doc.xpath("//c:val//c:v").each_with_index do |node, index|
            node.content = @collection.values[index]
          end

        elsif doc.xpath("//c:grouping").attr('val').value == 'stacked' # For Waterfall charts

          length = @collection.length
          doc.xpath("//c:ser").remove

          series_name = ['Fill', 'Base', 'Rise', 'Rise', 'Fall', 'Fall']
          color_fill_xml = [
            "<a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/>",
            "<a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill>",
            "<a:solidFill><a:schemeClr val=\"accent6\"/></a:solidFill>",
            "<a:solidFill><a:schemeClr val=\"accent6\"/></a:solidFill>",
            "<a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill>",
            "<a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill>"
          ]

          6.times do |i|
            series_temp = "<c:ser><c:idx val=\"#{i}\"/><c:order val=\"#{i}\"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>#{series_name[i]}</c:v></c:pt></c:strCache></c:strRef></c:tx><c:spPr>#{color_fill_xml[i]}</c:spPr><c:invertIfNegative val=\"0\"/><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val=\"1\"/></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"1\"/></c:numCache></c:numRef></c:val></c:ser>"
            doc.xpath("//c:barChart").first.add_child(series_temp)
          end

          doc.xpath("//c:ser").each do |series|
            column_idx = 0

            until series.xpath(".//c:cat").xpath(".//c:pt").length == length
              column_temp = "<c:pt idx=\"#{column_idx}\"><c:v>New Column</c:v></c:pt>"
              value_temp = "<c:pt idx=\"#{column_idx}\"><c:v>0.0</c:v></c:pt>"

              series.xpath(".//c:strCache").last.add_child(column_temp)
              series.xpath(".//c:numCache").last.add_child(value_temp)

              column_idx += 1
            end
          end

          output = etl_waterfall

          doc.xpath("//c:cat").xpath(".//c:v").each_with_index do |node, index|

            until index < @collection.length
              index -= @collection.length
            end
            node.content = @collection.keys[index]
          end

          doc.xpath("//c:val").xpath(".//c:v").each_with_index do |node, index|
            series = 0
            until index < output.values.first.length
              index -= output.values.first.length
              series += 1
            end
            node.content = output.values[series][index]
          end

        elsif doc.namespaces.include? 'xmlns:c' and doc.xpath("//c:barChart").any? # For Bar/Column Charts

          no_series = @collection.values.first.length
          doc.xpath("//c:ser").remove

          no_series.times do |i|
            series_temp = "<c:ser><c:idx val=\"#{i}\"/><c:order val=\"#{i}\"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>New Series</c:v></c:pt></c:strCache></c:strRef></c:tx><c:invertIfNegative val=\"0\"/><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val=\"1\"/></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"1\"/></c:numCache></c:numRef></c:val></c:ser>"
            doc.xpath("//c:barChart").first.add_child(series_temp)
          end

          length = @collection.length

          doc.xpath("//c:ser").each do |series|
            column_idx = 0

            until series.xpath(".//c:cat").xpath(".//c:pt").length == length
              column_temp = "<c:pt idx=\"#{column_idx}\"><c:v>New Column</c:v></c:pt>"
              value_temp = "<c:pt idx=\"#{column_idx}\"><c:v>0.0</c:v></c:pt>"

              series.xpath(".//c:strCache").last.add_child(column_temp)
              series.xpath(".//c:numCache").last.add_child(value_temp)

              column_idx += 1
            end
          end

          doc.xpath("//c:cat").xpath(".//c:v").each_with_index do |node, index|

            until index < @collection.length
              index -= @collection.length
            end

            node.content = @collection.keys[index]
          end

          doc.xpath("//c:val").xpath(".//c:v").each_with_index do |node, index|
            series = 0

            until index < @collection.length
              index -= @collection.length
              series += 1
            end

            node.content = @collection.values[index][series]
          end

        end

      end

    end

  end

  private

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
    # input << 0
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

    output
  end

end

end
