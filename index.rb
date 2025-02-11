require 'roo'
require 'caracal'
require 'securerandom'

def excel_to_docx(excel_file)
  unless File.exist?(excel_file)
    puts "Ошибка: файл #{excel_file} не найден!"
    return
  end

  xlsx = Roo::Excelx.new(excel_file)
  docx_file = "output_#{SecureRandom.hex(4)}.docx"  # Генерация имени файла

  Caracal::Document.save(docx_file) do |docx|
    headers = xlsx.row(1)

    docx.table [[headers] + xlsx.each_row_streaming(drop: 1).map {|row| row.map {|cell| cell&.value.to_s}}] do

      border_color '000000'
      border_size 8
      width 5000
    end
    
  end

  puts "Файл #{docx_file} успешно создан!"
end

# Запуск функции с передачей файла Excel
excel_to_docx('Test.xlsx')