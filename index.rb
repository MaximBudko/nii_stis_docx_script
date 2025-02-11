require 'roo'
require 'caracal'
require 'securerandom'

def excel_to_docx(excel_file)
  unless File.exist?(excel_file)
    puts "Ошибка: файл #{excel_file} не найден!"
    return
  end

  doc_name = 'Output.docx'
  workbook = Roo::Spreadsheet.open(excel_file)
  sheet = workbook.sheet(0)

  Caracal::Document.save(doc_name) do |doc|
    doc.table do |t|
      sheet.each_row_streaming do |row|
        t.row do |r|
          [row[0], row[4], row[5], row[6]].each do |cell|
            r.cell do |c|
              c.text cell.value.to_s
            end
          end
        end
      end
    end
  end


  puts "Файл успешно создан!"
end


# Запуск функции с передачей файла Excel
excel_to_docx('Test.xlsx')
