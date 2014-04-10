
require 'happymapper'

module ExcelXml
  class Cell
    include HappyMapper
    tag "Cell"
    element :data, String, tag: "Data", single: true
    attribute :index, Integer, tag: "Index"
    attribute :merge_down, Integer, tag: "MergeDown"
    attribute :merge_across, Integer, tag: "MergeAcross"
    alias_method :orig_merge_down, :merge_down
    alias_method :orig_merge_across, :merge_across
    def merge_down
      orig_merge_down || 0
    end
    def merge_across
      orig_merge_across || 0
    end
  end
end

