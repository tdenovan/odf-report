module ODFReport

module Parser
  
module MarkupParser
  
  def self.parse_formatting(text)
    text.strip!
    text.gsub!(/<strong>(.+?)<\/strong>/)  { "</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space=\"preserve\"> #{$1}</w:t><\/w:r><w:r><w:t>" }
    text
  end
    
end
  
end

end