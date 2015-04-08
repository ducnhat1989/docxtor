module Docxtor
  class Generator
    class << self
      def generate(template, &block)
        template_parser = TemplateParser.new(template)
        parts = template_parser.parts

        running_elements = RunningElementsBuilder.new(&block).elements
        reference = ReferenceBuilder.new(running_elements)
        document = Document::Builder.new(reference, &block)

        parts += document.reference.elements
        parts << document.reference
        parts << document
        Package::Builder.new(parts)
      end
    end
  end
end
