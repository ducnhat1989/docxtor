module Docxtor
  module Document
    class Builder
      attr_accessor :content, :reference

      def initialize(reference, &block)
        @reference = reference
        @content = render(&block)
      end

      def filename
        "word/document.xml"
      end

      def render(&block)
        xml = ::Builder::XmlMarkup.new

        xml.instruct! :xml, :version => "1.0", :encoding => "UTF-8", :standalone => "yes"
        xml.w :document, "xmlns:wpc" => "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
          "xmlns:mo" => "http://schemas.microsoft.com/office/mac/office/2008/main",
          "xmlns:mc" => "http://schemas.openxmlformats.org/markup-compatibility/2006",
          "xmlns:mv" => "urn:schemas-microsoft-com:mac:vml",
          "xmlns:o" => "urn:schemas-microsoft-com:office:office",
          "xmlns:r" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "xmlns:m" => "http://schemas.openxmlformats.org/officeDocument/2006/math",
          "xmlns:v" => "urn:schemas-microsoft-com:vml",
          "xmlns:wp14" => "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
          "xmlns:wp" => "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
          "xmlns:w10" => "urn:schemas-microsoft-com:office:word",
          "xmlns:w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
          "xmlns:w14" => "http://schemas.microsoft.com/office/word/2010/wordml",
          "xmlns:wpg" => "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
          "xmlns:wpi" => "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
          "xmlns:wne" => "http://schemas.microsoft.com/office/word/2006/wordml",
          "xmlns:wps" => "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
          "mc:Ignorable" => "w14 wp14" do
          xml.w :body do
            xml << Document::Root.new(reference, &block).render(::Builder::XmlMarkup.new)

            render_running_elements xml
          end
        end

        xml.target!
      end

      def render_running_elements xml
        xml.w :sectPr do |xml|
          reference.elements.each do |re|
            next if re.type == :image
            xml.w re.reference_name.to_sym, "r:id" => re.reference_id, "w:type" => "#{re.pages}"
          end
        end
      end
    end
  end
end
