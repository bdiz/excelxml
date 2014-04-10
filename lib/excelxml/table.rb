
require 'excelxml/row'

module ExcelXml
  class Table
    include HappyMapper
    tag "Table"
    has_many :rows, Row
    attribute :column_count, Integer, tag: "ExpandedColumnCount"
    attribute :row_count, Integer, tag: "ExpandedRowCount"
  end
end

