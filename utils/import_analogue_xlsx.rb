require 'roo'
require 'json'
require 'fileutils'

module JsonAnalogue

  COLUMN_COUNT = 33  

  def self.xlsx_convertor(xlsx_path)
    xlsx = Roo::Excelx.new(xlsx_path)
    
    headers = xlsx.row(2).map { |h| h.to_s.strip }

    unique_headers = []
    counts = Hash.new(0)

    headers.each do |h|
        counts[h] += 1
        if counts[h] > 1
            unique_headers << "#{h}_#{counts[h]}"
        else
            unique_headers << h
        end
    end

    data = []

    (xlsx.first_row + 2 .. xlsx.last_row).each do |i|
        row = Hash[[unique_headers, xlsx.row(i).map { |v| v.nil? ? "" : v }].transpose]
        data << row
    end

    return data
  end

  def self.xlsx_to_json(xlsx_path, json_path)
    data = xlsx_convertor(xlsx_path)
    FileUtils.mkdir_p(File.dirname(json_path))
    File.open(json_path, 'w') do |f|
      f.write(JSON.pretty_generate(data))
    end
  end

end
