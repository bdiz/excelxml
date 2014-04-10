# ExcelXml

ExcelXml can parse the data out of an Excel XML Spreadsheet 2003.

See the Usage section below to get an idea of how it works.

## Installation

Add this line to your application's Gemfile:

    gem 'excelxml'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install excelxml

## Usage

```ruby
require 'excelxml'

# First step is to inherit from a worksheet parser.
class PersonParser < ExcelXml::Worksheet::Parser
  NAME = /name/i
  CITY = /city/i
  STATE = /state/i
  AGE = /age/i

  class Person < Struct.new(:name, :city, :state, :age); end

  # Override the #mandator_columns() method to auto discover your worksheets header row. 
  # #is_header? could alternatively be overridden.
  def mandatory_columns
    [NAME, CITY]
  end

  # A rows accessor will be available to your parser so that you can iterate through worksheet data.
  def persons
    @persons ||= rows.collect do |fields| 
      Person.new(fields[NAME], fields[CITY], fields[STATE], fields[AGE])
    end
  end
end

my_parsers = ExcelXml::Workbook::Parser.new(File.read(workbook), worksheet_parsers: PersonParser) 

# You'll get a PersonParser instance for each worksheet that matched PersonParser#mandatory_columns
# or PersonParser#is_header?.
my_parsers[PersonParser].length                          # => 1
person_parser = my_parsers[PersonParser].first           # => #<PersonParser:0x00555555f42ab8>

person_parser.worksheet.name                             # => "PersonSheet"
person_parser.persons.each do |person| 
  puts "#{person.name} lives in #{person.city}."
end

# Worksheets that did not match a parsers #mandatory_columns or #is_header? end up here.
raw_worksheet = my_parsers.unidentified_worksheets.first # => #<ExcelXml::Worksheet:0x00555555f43841>
raw_worksheet.name                                       # => "NoHeaderSheet"

# ExcelXml::Worksheet#rows is a 2-dimensional array of string representing your worksheet.
raw_worksheet.rows.each do |raw_row| 
  raw_row.each {|raw_data| puts raw_data }
end
```

## Contributing

1. Fork it ( http://github.com/bdiz/excelxml/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
