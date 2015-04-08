module Docxtor
  class ReferenceBuilder
    attr_accessor :elements

    def initialize elements
      @elements = elements
      @elements.each_with_index { |el, i| el.reference_id = "rId#{i+1}" }
    end
    
    def add_element running_element
      running_element.reference_id = "rId#{elements.length+1}"
      @elements << running_element
    end
    
    def last_element
      @elements.last
    end

    def filename
      "word/_rels/document.xml.rels"
    end

    def content
      xml = ::Builder::XmlMarkup.new

      xml.instruct! :xml, :version => "1.0", :encoding=>"UTF-8", :standalone => "yes"
      xml.Relationships "xmlns" => "http://schemas.openxmlformats.org/package/2006/relationships" do |xml|
        elements.each do |element|
          xml.Relationship "Id" => element.reference_id,
                           "Target" => element.name,
                           "Type" => element.reference_type
        end
      end

      xml.target!
    end
  end
end
