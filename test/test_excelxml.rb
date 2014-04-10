$LOAD_PATH.unshift File.expand_path('../..', __FILE__)
require 'minitest_helper'

describe ExcelXml do

  workbook = fixture_file_path('workbook.xml')

  class PersonParser < ExcelXml::Worksheet::Parser
    NAME = /name/i
    BIRTH = /date\s+of\s+birth/i
    PROFESSION = /profession/i
    CITY = /city/i
    STATE = /state/i
    AGE = /age/i
    MANDATORY_COLUMNS = [NAME, CITY]

    class Person < Struct.new(:name, :birth, :profession, :city, :state, :age); end

    def mandatory_columns
      MANDATORY_COLUMNS
    end
#     def is_header? row_number, fields
#       row_number == 1
#     end
    def persons
      @persons ||= rows.collect do |fields| 
        Person.new(fields[NAME], fields[BIRTH], fields[PROFESSION], fields[CITY], fields[STATE], fields[AGE])
      end
    end
  end

  EXPECTED_NO_HEADER_SHEET_INFO = [
    ["1", "2", "3", "4", "5", nil, nil],
    ["6", "7", "7", "8", "9", "10", nil],
    %w(11 12 13 14 14 14 14),
    %w(11 15 16 14 14 14 14),
    %w(11 17 18 14 14 14 14),
    %w(19 19 19 19 19 19 19),
  ]
  EXPECTED_PERSON_INFO = [
    %w(jon jon jon CA CA 1),
    %w(jon jon jon CA CA 2),
    %w(jon jon jon CA CA 2),
    %w(13 13 13 CA CA 4),
    ["13", "13", "13", "CA", "CA", nil],
    ["ben", "may", "programmer", "San Diego", "CA", "6"]
  ]

  it "it can parse a workbook" do
    workbook_parser = ExcelXml::Workbook::Parser.new(File.read(workbook), worksheet_parsers: PersonParser) 

    workbook_parser.unidentified_worksheets.length.must_equal 1
    workbook_parser.unidentified_worksheets.first.name.must_equal "NoHeaderSheet"
    workbook_parser.unidentified_worksheets.first.rows.must_equal EXPECTED_NO_HEADER_SHEET_INFO

    workbook_parser[PersonParser].length.must_equal 1
    person_parser = workbook_parser[PersonParser].first
    
    person_parser.worksheet.name.must_equal "PersonSheet"
    person_parser.persons.collect {|person| person.to_a }.must_equal EXPECTED_PERSON_INFO
  end


end
