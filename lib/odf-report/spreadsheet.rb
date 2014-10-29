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

  end

  def replace!(doc, filename, row = nil)

    if /worksheets/ === filename

      doc.xpath("//xmlns:row").remove

      rows = @collection.length + 1
      cols = @series.length + 1

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
            c_node = "<c r=\"#{(col + 65).chr}#{row + 1}\"><v>#{val}</v></c>"
            doc.xpath("//xmlns:row")[row].add_child(c_node)

          end

        end

      end

    elsif /tables/ === filename

      rows = @collection.length + 1
      cols = @series.length + 1

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

      doc.xpath("//xmlns:si").last.remove until doc.xpath("//xmlns:si").length == 1

      @series.each do |name|
        node = "<si><t>#{name}</t></si>"
        doc.xpath("//xmlns:sst").last.add_child(node)
      end

      @collection.keys.each do |name|
        node = "<si><t>#{name}</t></si>"
        doc.xpath("//xmlns:sst").last.add_child(node)
      end

    end

    # debugger
    # 'hek'

  end

  end

end
