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
    
    xlsx.each_row_streaming do |row|
      text = row.map { |cell| cell&.value.to_s }.join(' | ')  # Форматируем строку
      docx.p(text) unless text.strip.empty?  # Добавляем строку как параграф в документ
    end
    
  end

  puts "Файл #{docx_file} успешно создан!"
end

# Запуск функции с передачей файла Excel
excel_to_docx('Test.xlsx')