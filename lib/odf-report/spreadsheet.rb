module ODFReport

class Spreadsheet
  include Nested

  @@id_target = {}
  @@name_id = {}
  @sheet_id = {}

  def initialize(opts)

    @name             = opts[:name]
    @collection       = opts[:collection]

    @series           = opts[:series] || nil
    @title            = opts[:title] || nil
    @type             = opts[:type] || nil

  end

  def replace!(doc, filename, row = nil)

    case @type

    when 'pie', 'doughnut' # For Pie Charts

      @series = [@series] unless @series.is_a? Array
      @collection.each { |k, v| @collection[k] = [v] } unless @collection.values.first.is_a? Array

    when 'waterfall' # For Waterfall Charts

      etl_waterfall unless @series == ['Fill', 'Base', 'Rise+', 'Rise-', 'Fall+', 'Fall-']

    when 'bar', 'column', 'line' # For Bar/Column/Line Charts

      if @series.length < @collection.values.first.length # If number of series names doesn't match the collection values

        @series << rand(65..91).chr until @series.length == @collection.values.first.length

      elsif @series.length < @collection.values.first.length

        @series.pop until @series.length == @collection.values.first.length

      end

    end

    rows = @collection.length + 1
    cols = @series.length + 1

    if /worksheets/ === filename

      doc.xpath("//xmlns:row").remove
      doc.xpath("//xmlns:dimension").first['ref'] = "A1:#{(cols + 64).chr}#{rows}"

      rows.times do |row|
        node = "<row r=\"#{row + 1}\" spans=\"1:4\">"
        doc.xpath("//xmlns:sheetData").first.add_child(node)
      end

      si = 0

      rows.times do |row|

        cols.times do |col|

          if row == 0 or col == 0

            s_node = "<c r=\"#{(col + 65).chr}#{row + 1}\" t=\"s\"><v>#{si}</v></c>"
            doc.xpath("//xmlns:row")[row].add_child(s_node)
            si += 1

          else

            val = @collection.values[row - 1][col - 1]
            val = 0 if @collection.values[row - 1][col - 1].nil?
            c_node = "<c r=\"#{(col + 65).chr}#{row + 1}\"><v>#{val}</v></c>"
            doc.xpath("//xmlns:row")[row].add_child(c_node)

          end

        end

      end

    elsif /tables/ === filename

      doc.xpath("//xmlns:table").first['ref'] = "A1:#{(cols + 64).chr}#{rows}"
      doc.xpath("//xmlns:tableColumns").first['count'] = @series.length + 1

      node = doc.xpath("//xmlns:tableColumn").first
      doc.xpath("//xmlns:tableColumn").remove
      doc.xpath("//xmlns:tableColumns").first.add_child(node)

      @series.each do |name|

        last_id = doc.xpath("//xmlns:tableColumn").last['id']
        id = (last_id.to_i + 1).to_s
        node = "<tableColumn id=\"#{id}\" name=\"#{name}\"/>"

        doc.xpath("//xmlns:tableColumns").last.add_child(node)

      end

    elsif /sharedStrings/ === filename

      doc.xpath("//xmlns:si").remove

      first_node = "<si><t xml:space=\"preserve\"> </t></si>"
      doc.xpath("//xmlns:sst").last.add_child(first_node)

      @series.each do |name|
        node = "<si><t>#{name}</t></si>"
        doc.xpath("//xmlns:sst").last.add_child(node)
      end

      @collection.keys.each do |name|
        node = "<si><t>#{name}</t></si>"
        doc.xpath("//xmlns:sst").last.add_child(node)
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

    @series = ['Fill', 'Base', 'Rise+', 'Rise-', 'Fall+', 'Fall-']

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
