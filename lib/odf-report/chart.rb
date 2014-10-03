module ODFReport

class Chart
  include Nested

  def initialize(opts)
    @name             = opts[:name]
    @collection       = opts[:collection]

    @series_name      = opts[:series_name] || ''
    @chart_type       = opts[:chart_type] || ''

    $id_target = {}
    $name_id = {}

  end

  def replace!(doc, row = nil)

    if doc.namespaces.include? 'xmlns' and doc.xpath("//xmlns:Relationship").any? # Look into _rels/document.xml.rels

      doc.xpath("//xmlns:Relationship").each do |rel|
        $id_target[rel.attr('Id')] = rel.attr('Target')
      end

    elsif doc.namespaces.include? 'xmlns:w' and doc.xpath("//w:drawing").any? # Look into word/document.xml

      $name_id[@name] = doc.xpath("//w:drawing//wp:docPr[@title='#{@name}']/following-sibling::*").xpath(".//c:chart", {'c' => "http://schemas.openxmlformats.org/drawingml/2006/chart"}).attr('id').value

    elsif doc.namespaces.include? 'xmlns:c' # Look through chart1.xml

      # ****************************************************************************************************
      $count ||= 2

      if $id_target[$name_id[@name]] == "charts/chart#{$count}.xml"
        $count -= 1
        # The line above NEEDS to be changed
      # ****************************************************************************************************

        # doc.xpath("//c:v") is structured as follows:
        # - First(index = 0) is always the Series Type
        # - Next(index 1..4) are the types in the series
        # - Last(index 5..8) are the datas for the corresponding types
        # Hence why the formula below looks very weird

        length = (doc.xpath("//c:v").length - 1) / 2

        doc.xpath("//c:v").each_with_index do |node, index|

          index -= 1

          if index < 0
            # doc.at("//text()[.='#{node.text}']").content = 'Alphabet' # Series Name should be in the field
          elsif index < length
            doc.at("//text()[.='#{node.text}']").content = @collection.keys[index]
          else
            doc.at("//text()[.='#{node.text}']").content = @collection.values[index - length]
          end

        end

      end

    end

  end

end

end
