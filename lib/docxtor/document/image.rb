module Docxtor
  module Document
    class Image < Element
      AUTO_WIDTH = 5
      AUTO_HEIGHT = 5

      def after_initialize(link, *args)
        @rId = @reference.last_element.reference_id
        @num = @rId
        style = args.shift || {}
        @width = (style[:width] || AUTO_WIDTH) * 9144000 / 96
        @height = (style[:height] || AUTO_HEIGHT) * 9144000 / 96
      end

      def render xml
        xml.w :p do
          xml.w :r do
            xml.w :rPr do
              xml.w :noProof
            end
            xml.w :drawing do
              xml.wp :inline, "distT" => "0", "distB" => "0", "distL" => "0", "distR" => "0" do
                xml.wp :extent, "cx" => @height, "cy" => @width
                xml.wp :effectExtent, "l" => "0", "t" => "0", "r" => "0", "b" => "0"
                xml.wp :docPr, "id" => @num, "name" => "Picture #{@num}"
                xml.wp :cNvGraphicFramePr do
                  xml.a :graphicFrameLocks, "xmlns:a" => "http://schemas.openxmlformats.org/drawingml/2006/main", "noChangeAspect" => "1"
                end
                xml.a :graphic, "xmlns:a" => "http://schemas.openxmlformats.org/drawingml/2006/main" do
                  xml.a :graphicData, "uri" => "http://schemas.openxmlformats.org/drawingml/2006/picture" do
                    xml.pic :pic, "xmlns:pic" => "http://schemas.openxmlformats.org/drawingml/2006/picture" do
                      xml.pic :nvPicPr do
                        xml.pic :cNvPr, "id" => @num, "name" => "Picture #{@num}"
                        xml.pic :cNvPicPr do
                          xml.a :picLocks, "noChangeAspect" => "1", "noChangeArrowheads" => "1"
                        end
                      end
                      xml.pic :blipFill do
                        xml.a :blip, "r:embed" => "#{@rId}"
                        xml.a :srcRect
                        xml.a :stretch do
                          xml.a :fillRect
                        end
                      end
                      xml.pic :spPr, "bwMode" => "auto" do
                        xml.a :xfrm do
                          xml.a :off, "x" => "0", "y" => "0"
                          xml.a :ext, "cx" => @height, "cy" => @width
                        end
                        xml.a :prstGeom, "prst" => "rect" do
                          xml.a :avLst
                        end
                        xml.a :noFill
                        xml.a :ln do
                          xml.a :noFill
                        end
                      end
                    end
                  end
                end
              end
            end
          end
        end
      end
    end
  end
end