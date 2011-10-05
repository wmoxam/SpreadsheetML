require "rexml/document"

class SpreadsheetML
  attr_accessor :worksheets

  def initialize(stream_or_string)
    @worksheets = []

    return if stream_or_string.nil?

    doc = REXML::Document.new stream_or_string

    doc.elements.each("Workbook/Worksheet") { |worksheet| @worksheets << Worksheet.new(worksheet) }
  end

  class Worksheet
    attr_accessor :name, :tables

    def initialize(xml)
      @name = xml.attributes["ss:Name"]
      @tables = xml.elements.collect("Table") { |table| Table.new(table) }
    end
  end

  class Table
    attr_accessor :rows

    def initialize(xml)
      @rows = xml.elements.collect("Row") { |row| Row.new(row) }
    end
  end

  class Row
    attr_accessor :cells

    def initialize(xml)
      @cells = xml.elements.collect("Cell") { |cell| Cell.new(cell) }
    end
  end

  class Cell
    attr_accessor :text

    # ignoring style
    def initialize(xml)
      @text = ''
      xml.each_element_with_text {|e| @text << e.text.strip }
    end

    def to_s
      @text
    end
  end
end
