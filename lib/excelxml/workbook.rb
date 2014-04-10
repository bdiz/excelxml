
require 'excelxml/worksheet'

module ExcelXml

  class Workbook

    include HappyMapper
    tag "Workbook"
    has_many :worksheets, Worksheet

    class Parser
      attr_reader :unidentified_worksheets
      def initialize workbook_xml, opts={}
        only_these_worksheets = [opts.delete(:only_these_worksheets)].flatten.compact if opts[:only_these_worksheets]
        @worksheet_parser_classes = [opts.delete(:worksheet_parsers)].flatten.compact
        @worksheet_parser_hash = @worksheet_parser_classes.each_with_object({}) {|wspc, hsh| hsh[wspc] = [] }
        raise ArgumentError, "unknown options #{opts.keys.inspect}" unless opts.empty?
        @unidentified_worksheets = []
        ExcelXml::Workbook.parse(workbook_xml, single: true).worksheets.each do |worksheet|
          next if only_these_worksheets and !only_these_worksheets.include?(worksheet.name)
          worksheet_identified = false
          worksheet.rows.each_with_index do |row, row_idx|
            worksheet_identifiers.each do |wsp|
              if wsp.is_header? (row_idx+1), row
                add_worksheet_parser(wsp.class, worksheet, row_idx)
                worksheet_identified = true
                break
              end
            end
            break if worksheet_identified
          end unless @worksheet_parser_classes.empty?
          @unidentified_worksheets << worksheet unless worksheet_identified
        end
      end

      def [] worksheet_parser_class
        @worksheet_parser_hash[worksheet_parser_class]
      end

      private

      def add_worksheet_parser worksheet_parser_class, worksheet, header_row_idx
        @worksheet_parser_hash[worksheet_parser_class] << worksheet_parser_class.new(worksheet)
        @worksheet_parser_hash[worksheet_parser_class].last.header_row_index = header_row_idx
      end

      def worksheet_identifiers 
        @worksheet_identifiers ||= @worksheet_parser_classes.collect {|wspc| wspc.new(nil) }
      end

    end

  end
end

