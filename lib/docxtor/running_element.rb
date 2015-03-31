module Docxtor
  class RunningElement
    attr_accessor :pages, :type, :reference_id

    def initialize type, num, contents, options = {}
      @type = type
      @contents = contents
      @align = options[:align]
      @pages = options[:pages] || :default
      @num = num
    end

    def reference_name
      "#{type}Reference"
    end

    def name
      "#{type}#{@num}.xml"
    end

    def filename
      "word/#{name}"
    end

    def reference_type
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/#{type}"
    end

    def content
      xml = ::Builder::XmlMarkup.new

      if @type == :header
        xml.instruct! :xml, :version => "1.0", :encoding => "UTF-8", :standalone => "yes"
        xml.w :hdr,
          "xmlns:wpc" => "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
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
          "mc:Ignorable" => "w14 wp14" do |xml|
          xml.w :p do |xml|
            xml.w :pPr do |xml|
              xml.w :jc, "w:val" => "#{@align}" if @align
              xml.w :pStyle, "w:val" => "Header"
              xml.w :proofErr, "w:type" => "spellStart"
              xml.w :proofErr, "w:type" => "gramStart"
              xml.w :r do |xml|
                xml.w :t, @contents
              end
              xml.w :proofErr, "w:type" => "spellEnd"
              xml.w :proofErr, "w:type" => "gramEnd"
            end
          end
        end
      else
        xml.instruct! :xml, :version => "1.0", :encoding=>"UTF-8", :standalone => "yes"
        xml.w :ftr, "xmlns:o" => "urn:schemas-microsoft-com:office:office",
        "xmlns:r" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:v" => "urn:schemas-microsoft-com:vml",
        "xmlns:w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:w10" => "urn:schemas-microsoft-com:office:word",
        "xmlns:wp" => "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" do |xml|

          xml.w :p do |xml|
            xml.w :pPr do |xml|
              xml.w :jc, "w:val" => "#{@align}" if @align
            end
            if @contents == :pagenum
              xml.w :r do |xml|
                xml.w :fldChar, "w:fldCharType" => "begin"
              end
              xml.w :r do |xml|
                xml.w :instrText, "PAGE"
              end
              xml.w :r do |xml|
                xml.w :fldChar, "w:fldCharType" => "separate"
              end
              xml.w :r do |xml|
                xml.w :t, "i"
              end
              xml.w :r do |xml|
                xml.w :fldChar, "w:fldCharType" => "end"
              end
            else
              xml.w :r do |xml|
                xml.w :t, @contents
              end
            end
          end
        end
      end
    end
  end
end
