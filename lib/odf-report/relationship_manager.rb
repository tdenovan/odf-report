module ODFReport

  # This class is responsible for managing relations in the document.xml.rels file
  # It helps create new relationships for new images, etc
  # This class is effectively a singleton
  class RelationshipManager

    RELATIONSHIP_TYPES = {
      image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      slide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    }
    RELATIONSHIP_FILES = {
      doc: "word/_rels/document.xml.rels",
      ppt: "ppt/_rels/presentation.xml.rels"
    }
    TARGET_EXTENSIONS = {
      slide: 'xml',
      image: 'png'
    }
    TARGET_DIRECTORIES = {
      slide: 'slides',
      image: 'media'
    }
    ID_PREFIX = 'rId'

    def initialize(file_instance, file_type)
      @file_type = file_type
      @current_id = 1000 # This is the next id that has not yet been assigned to a relationship element. Ideally this should be read this from the appropriate file
      @relationships = []
      @new_relationships = []
      @file = file_instance
      @created_targets = {}
    end

    # Parses a relationships xml node and records all the relationships
    def parse_relationships(content)
      if content.xpath("//xmlns:Relationship").any? # Looking through word/_rels/document.xml.rels
        content.xpath("//xmlns:Relationship").each do |rel|
          @relationships << {
            id: rel.attr('Id'),
            target: rel.attr('Target'),
            type: rel.attr('Type')
          }
        end
      end
    end
    
    # Gets a relation by its target
    def get_relationship_by_target(target)
      @relationships.select { |relationship| relationship[:target] == target }.first
    end

    # Get relationship by id
    def get_relationship(id)
      @relationships.select { |relationship| relationship[:id] == id }.first
    end

    # Gets all relationships by type
    def get_relationships(type)
      @relationships.select { |relationship| relationship[:type] == type }
    end

    # Gets or create a target path
    def create_target_and_id(type, path)
      @current_id += 1
      id = "#{ID_PREFIX}#{@current_id}"
      target = "#{TARGET_DIRECTORIES[type]}/#{type.to_s}#{id}.#{TARGET_EXTENSIONS[type]}"
      path = target if path.nil?
      @created_targets[path] = { target: target, id: id }
      return target, id
    end

    # Creates a new relationship and assigns it an id and a path
    def new_relationship(type, path = nil)
      raise "Unknown relationship type #{type}" if RELATIONSHIP_TYPES[type].nil?
      return @created_targets[path][:id], @created_targets[path][:target] unless path.nil? or @created_targets[path].nil? # Check if it exists already

      target, id = create_target_and_id(type, path)

      @new_relationships << {
        id: id,
        target: target,
        type: RELATIONSHIP_TYPES[type]
      }
      return id, target
    end

    # Writes new relationships to the relationships document
    def write_new_relationships
      data = @file.read_file(RELATIONSHIP_FILES[@file_type])
      doc = Nokogiri::XML(data)
      relationships = doc.xpath("//xmlns:Relationships").first
      return if relationships.nil?

      @new_relationships.each do |relationship_hash|
        new_node = Nokogiri::XML::Node.new "Relationship", doc
        new_node.set_attribute('Target', relationship_hash[:target])
        new_node.set_attribute('Type', relationship_hash[:type])
        new_node.set_attribute('Id', relationship_hash[:id])
        relationships.add_child new_node
      end

      data.replace(doc.to_xml(:save_with => Nokogiri::XML::Node::SaveOptions::AS_XML))
      @file.output_stream.put_next_entry(RELATIONSHIP_FILES[@file_type])
      @file.output_stream.write data
    end

  end # ImageManager class
end # ODFReport module
