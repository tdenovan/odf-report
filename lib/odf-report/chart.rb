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

    elsif doc.namespaces.include? 'xmlns:c' # Look through chart1.xml

      if filename.include? @@id_target[@@name_id[@name]]

        if doc.namespaces.include? 'xmlns:c' and doc.xpath("//c:pieChart").any?

          doc.xpath("//c:cat//c:v").each_with_index do |node, index|
            node.content = @collection.keys[index]
          end

          doc.xpath("//c:val//c:v").each_with_index do |node, index|
            node.content = @collection.values[index]
          end

        elsif doc.xpath("//c:grouping").attr('val').value.nil? == false # For waterfall charts

          output = rearrange

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

        elsif doc.namespaces.include? 'xmlns:c' and doc.xpath("//c:barChart").any?

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

  def rearrange
    input = @collection.values
    output = {
      :fill => [],
      :base => [],
      :more => [],
      :less => []
    }
    input << 0
    input.each_with_index do |num, index|
      if index == 0
        output[:fill] << 0
        output[:base] << num
        output[:more] << 0
        output[:less] << 0
      elsif index == input.length - 1
        output[:base] << output[:fill].last + output[:more].last if output[:more].last > 0
        output[:base] << output[:fill].last if output[:less].last > 0
        output[:fill] << 0
        output[:more] << 0
        output[:less] << 0
      elsif num >= 0
        output[:fill] << (output[:fill].last + output[:base].last + output[:more].last)
        output[:base] << 0
        output[:more] << num
        output[:less] << 0
      elsif num < 0
        output[:fill] << (output[:fill].last + output[:base].last + output[:more].last + num)
        output[:base] << 0
        output[:more] << 0
        output[:less] << - num
      end
    end

    puts output
    output
  end

end

end
