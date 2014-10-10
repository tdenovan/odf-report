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

          doc.xpath("//c:cat//c:v").each_with_index do |node, index|
            node.content = @collection.keys[index]
          end

          doc.xpath("//c:val//c:v").each_with_index do |node, index|
            node.content = @collection.values[index]
          end

        elsif doc.xpath("//c:grouping").attr('val').value.nil? == false # For Waterfall charts

          output = rearrange_waterfall

          doc.xpath("//c:cat").xpath(".//c:v").each_with_index do |node, index|

            until index < @collection.length
              index -= @collection.length
            end
            @collection['Finish'] = nil
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

  def rearrange_waterfall
    input = @collection.values
    output = {
      'fill' => [],
      'base' => [],
      'pro+' => [],
      'pro-' => [],
      'con+' => [],
      'con-' => []
    }
    input << 0
    sum = 0

    input.each_with_index do |num, index|

      if index == 0
        output['base'] << num
      elsif index == input.length - 1
        output['base'] << sum
      elsif num > 0 and sum > 0
        output['fill'] << sum
        output['pro+'] << num
      elsif num > 0 and sum < 0 and num < sum
        output['fill'] << sum - num
        output['pro-'] << - num
      elsif num > 0 and sum < 0 and num > sum
        output['pro+'] << sum + num
        output['pro-'] << sum
      elsif num < 0 and sum < 0
        output['fill'] << sum
        output['con-'] << num
      elsif num < 0 and sum > 0 and sum + num > 0
        output['fill'] << sum + num
        output['con+'] << - num
      elsif num < 0 and sum > 0 and sum + num < 0
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
