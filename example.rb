#!/usr/bin/env ruby

require 'spreadsheetml'
require 'csv'

# Converts SpreadsheetML to CSV

SpreadsheetML.new(ARGF.read).worksheets.each do |worksheet|
  puts "Title: #{worksheet.name}" unless worksheet.name.nil?

  worksheet.tables.each do |table|
    table.rows.each do |row|
      puts CSV::Row.new(row.cells, row.cells).to_s
    end
  end
  puts "\n"
end
