require 'rubyXL'
require 'zip'
require 'nokogiri'
require 'fileutils'

# Словарь замен единиц измерения
UNIT_MAP = {
  'n' => 'нФ',
  'u' => 'мкФ',
  'm' => 'МОм',
  'k' => 'кОм',
  'p' => 'пФ'
}

# Пути к файлам
docx_path = "shablon_pr.docx"
new_docx_path = "shablon_pr_updated.docx"
xlsx_path = "Test.xlsx"

# Функция обработки характеристик
def parse_characteristics(value, tolerance)
  return "" if value.nil? || value.strip.empty?

  value = value.gsub(",", ".") # Заменяем запятую на точку

  unit = value[/[a-zA-Z]+/] # Извлекаем единицу измерения
  number = value[/\d+(\.\d+)?/] # Извлекаем число
  dnp = value.include?("DNP") ? " DNP" : ""

  unit = UNIT_MAP[unit] || unit # Подставляем русское обозначение

  formatted_value = number ? "#{number} #{unit}" : value
  formatted_value += "±#{tolerance}" unless tolerance.nil? || tolerance.empty?

  formatted_value.gsub(".", ",") + dnp
end

# Функция форматирования номеров
def format_numbers(numbers)
  return numbers.first if numbers.size == 1

  sorted_numbers = numbers.sort_by { |num| num[/\d+/].to_i rescue num }
  ranges = []
  temp_range = [sorted_numbers.first]

  sorted_numbers.each_cons(2) do |prev, curr|
    prev_num = prev[/\d+/].to_i rescue prev
    curr_num = curr[/\d+/].to_i rescue curr

    if curr_num == prev_num + 1
      temp_range << curr
    else
      ranges << (temp_range.size > 2 ? "#{temp_range.first}-#{temp_range.last}" : temp_range.join(","))
      temp_range = [curr]
    end
  end

  ranges << (temp_range.size > 2 ? "#{temp_range.first}-#{temp_range.last}" : temp_range.join(",")) unless temp_range.empty?
  ranges.join(", ")
end

# Открываем Excel
xlsx = RubyXL::Parser.parse(xlsx_path)
sheet = xlsx[0] # Берем первый лист

# Читаем и обрабатываем данные из Excel
data = []
last_value = nil
count = 1
current_numbers = []

# Пропускаем первую строку с заголовками
sheet.each_with_index do |row, index|
  next if index == 0  # Пропускаем первую строку (заголовки)

  current_value = row[1]&.value.to_s.strip # 2-я колонка
  current_number = row[0]&.value.to_s.strip # 1-я колонка

  next if current_value.empty?

  description = "#{row[4]&.value.to_s.strip} #{row[5]&.value.to_s.strip}"
  characteristics = parse_characteristics(row[2]&.value.to_s.strip, row[6]&.value.to_s.strip)

  if current_value == last_value
    count += 1
    current_numbers << current_number
    data.last[2] = count.to_s
  else
    data.last[0] = format_numbers(current_numbers) unless current_numbers.empty?
    count = 1
    current_numbers = [current_number]
    data << [
      current_number,
      description,   # Описание
      "1",           # Количество
      characteristics # Характеристики
    ]
  end

  last_value = current_value
end

data.last[0] = format_numbers(current_numbers) unless current_numbers.empty?

# Работа с Word
FileUtils.cp(docx_path, new_docx_path)
Zip::File.open(new_docx_path) do |zip|
  document_xml = zip.find_entry("word/document.xml")

  if document_xml
    xml_content = document_xml.get_input_stream.read
    doc = Nokogiri::XML(xml_content)

    # Сохраняем исходный XML для отладки
    File.write("before_edit.xml", doc.to_xml)

    # Ищем таблицы
    tables = doc.xpath("//w:tbl", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

    puts "🔹 Найдено таблиц: #{tables.size}"
    
    tables.each do |table|
      last_row = table.xpath(".//w:tr").last # Последняя строка
      puts "🔹 Найдено строк в таблице: #{table.xpath('.//w:tr').size}"

      data.each do |row_data|
        new_row = Nokogiri::XML::Node.new("w:tr", doc) # Создаем строку

        row_data.each do |value|
          cell = Nokogiri::XML::Node.new("w:tc", doc) # Создаем ячейку
          paragraph = Nokogiri::XML::Node.new("w:p", doc) # Создаем параграф
          run = Nokogiri::XML::Node.new("w:r", doc) # Создаем run (контейнер для текста)
          text_node = Nokogiri::XML::Node.new("w:t", doc) # Создаем текстовый узел

          text_node.content = value.empty? ? "[ПУСТО]" : value

          # Применение стиля шрифта GOST type A, размер 14, курсив
          run_properties = Nokogiri::XML::Node.new("w:rPr", doc)
          font = Nokogiri::XML::Node.new("w:rFonts", doc)
          font['w:ascii'] = "GOST Type A"
          font['w:hAnsi'] = "GOST Type A"
          font['w:eastAsia'] = "GOST Type A"
          font['w:cs'] = "GOST Type A"
          run_properties.add_child(font)

          size = Nokogiri::XML::Node.new("w:sz", doc)
          size['w:val'] = "28"  # Размер шрифта 14 (в половинных пунктах)
          run_properties.add_child(size)

          italic = Nokogiri::XML::Node.new("w:i", doc) # Курсив
          run_properties.add_child(italic)

          run.add_child(run_properties)
          run.add_child(text_node)
          paragraph.add_child(run)
          cell.add_child(paragraph)
          new_row.add_child(cell)
        end

        puts "✅ Добавлена строка: #{row_data.inspect}" # Лог добавления строки
        table.add_child(new_row) # Вставляем строку в таблицу
      end
    end

    # Сохраняем измененный XML для отладки
    File.write("after_edit.xml", doc.to_xml)

    # Записываем изменения обратно в docx
    zip.get_output_stream("word/document.xml") { |f| f.write(doc.to_xml) }
  end
end

puts "✅ Данные успешно добавлены в shablon_pr_updated.docx"
