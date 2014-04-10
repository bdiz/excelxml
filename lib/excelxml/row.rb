
require 'excelxml/cell'

module ExcelXml
  class Row
    include HappyMapper
    tag "Row"
    has_many :cells, Cell
    attribute :index, Integer, tag: "Index"
  end
end

