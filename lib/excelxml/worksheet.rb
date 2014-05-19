
require 'excelxml/table'

module ExcelXml

  class Worksheet

    include HappyMapper
    tag "Worksheet"
    attribute :name, String, tag: "Name"
    element :table, Table, :single => true

    def rows
      @grid ||= begin
        grid = Array.new(table.row_count) { Array.new(table.column_count) }
        grid_row_idx = -1
        table.rows.each_with_index do |row, row_idx|
          grid_row_idx = row.index ? row.index - 1 : grid_row_idx + 1
          grid_col_idx = -1
          row.cells.each_with_index do |cell, cell_idx|
            grid_col_idx = cell.index ?  cell.index - 1 : grid_col_idx + 1
            (0..cell.merge_down).each do |down|
              (0..cell.merge_across).each do |across|
                grid[grid_row_idx+down][grid_col_idx+across] = Field.new(cell.data || "")
              end
            end
            grid_col_idx += cell.merge_across
          end
        end
        grid
      end
    end

    class Parser

      attr_reader :worksheet
      attr_accessor :header_row_index

      def initialize worksheet
        @worksheet = worksheet
      end

      def rows
        @rows ||= begin
          @worksheet.rows[(header_row_index+1)..-1].collect do |fields|
            Fields.new(self, fields)
          end
        end
      end

      def has_header? header
        index_to_header_map.find {|e| e.match header }
      end

      def index_to_header_map
        @index_to_header_map ||= @worksheet.rows[header_row_index].collect {|h| h ? h.strip : nil }
      end

      ########################################
      # Override methods
      ########################################

      def mandatory_columns
        []
      end

      def is_header? row_number, fields
        return true if mandatory_columns.all? {|mc| fields.compact.any? {|f| f.match mc } }
        return false
      end

    end
  end

  class Fields < Array
    def initialize parser, raw_fields
      super(raw_fields)
      @parser = parser
    end
    def has_header? header
      @parser.has_header? header
    end
    def [] key
      if idx = @parser.instance_eval { index_to_header_map.find_index {|e| e.match key } }
        self[idx] = super(idx) || Field.new
      end
    end
  end

  class Field < String
    def fixnum? 
      return Integer(self)
    rescue
      return false
    end
    def content?
      return !self.empty?
    end
    def string? 
      return (self.content? and !self.fixnum?)
    end
    def to_nil
      self.empty? ? nil : self
    end
  end

end
