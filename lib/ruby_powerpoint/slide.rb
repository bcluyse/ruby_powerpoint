require 'zip/zipfilesystem'
require 'nokogiri'

module RubyPowerpoint
  class RubyPowerpoint::Slide

    attr_reader :presentation,
                :slide_number,
                :slide_number,
                :slide_file_name

    def initialize presentation, slide_xml_path
      @presentation = presentation
      @slide_xml_path = slide_xml_path
      @slide_number = extract_slide_number_from_path slide_xml_path
      @slide_notes_xml_path = "ppt/notesSlides/notesSlide#{@slide_number}.xml"
      @slide_file_name = extract_slide_file_name_from_path slide_xml_path

      parse_slide
      parse_slide_notes
      parse_relation
    end

    def parse_slide 
      slide_doc = @presentation.files.file.open @slide_xml_path
      @slide_xml = Nokogiri::XML::Document.parse slide_doc
    end

    def parse_slide_notes
      slide_notes_doc = @presentation.files.file.open @slide_notes_xml_path rescue nil
      if slide_notes_doc
        @slide_notes_xml = Nokogiri::XML::Document.parse(slide_notes_doc)
      end 
    end

    def parse_relation
      @relation_xml_path = "ppt/slides/_rels/#{@slide_file_name}.rels"
      if @presentation.files.file.exist? @relation_xml_path
        relation_doc = @presentation.files.file.open @relation_xml_path
        @relation_xml = Nokogiri::XML::Document.parse relation_doc
      end
    end

    def content
      content_elements @slide_xml
    end

    def notes_content
      if(@slide_notes_xml)
        content_elements = content_elements(@slide_notes_xml)
        # Cut out page number
        content_elements.pop
        # Join the notes together
        content_elements.join if content_elements.length > 0 

      else 
        return nil
      end

    end

    def change_title(new_title, old_title)
      if(title == old_title)

        # Find the title
        temp = nil
        @slide_xml.xpath('//p:sp').each do |node|
          if(element_is_title(node))
            node.xpath('//a:t').each do |attempt|
              if(attempt.content == old_title)
                puts attempt.content
                attempt.content = new_title
              end
            end
          end
        end

        # Write to file
        @presentation.files.get_output_stream(@slide_xml_path) { |f| f.puts @slide_xml }
        # File.write(@slide_xml_path, @slide_xml)
       


        return
      end
    end

    
    def title
      title_elements = title_elements(@slide_xml)
      title_elements.join(" ") if title_elements.length > 0
    end

    def images
      image_elements(@relation_xml)
        .map.each do |node|
          @presentation.files.file.open(
            node['Target'].gsub('..', 'ppt'))
        end
    end
   
    def slide_num
      @slide_xml_path.match(/slide([0-9]*)\.xml$/)[1].to_i
    end
 
    private

    def extract_slide_number_from_path path
      path.gsub('ppt/slides/slide', '').gsub('.xml', '').to_i
    end

    def extract_slide_file_name_from_path path
      path.gsub('ppt/slides/', '')
    end

    def title_elements(xml)
      shape_elements(xml).select{ |shape| element_is_title(shape) }
    end
    
    def content_elements(xml)
      xml.xpath('//a:t').collect{ |node| node.text }
    end

    def image_elements(xml)
      xml.css('Relationship').select{ |node| element_is_image(node) }
    end    

    def shape_elements(xml)
      xml.xpath('//p:sp')
    end    
  
    def element_is_title(shape)
      shape.xpath('.//p:nvSpPr/p:nvPr/p:ph').select{ |prop| prop['type'] == 'title' || prop['type'] == 'ctrTitle' }.length > 0
    end

    def element_is_image(node)
      node['Type'].include? 'image' 
    end
  end
end
