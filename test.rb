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

# Словарь соответствий типов компонентов
CATEGORY_MAP = {
  'G' => 'Генераторы',
  'C' => 'Конденсаторы',
  'D' => 'Микросхемы',
  'DA' => 'Микросхемы аналоговые',
  'F' => 'Предохранители',
  'HL' => 'Индикаторы',
  'K' => 'Реле',
  'L' => 'Дросили',
  'R' => 'Резисторы',
  'SB' => 'Кнопки тактовые',
  'U' => 'Модули',
  'VD' => 'Диоды',
  'VT' => 'Транзисторы',
  'X' => 'Соединители'
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
last_category = nil  # Переменная для хранения последней категории

# Пропускаем первую строку с заголовками
sheet.each_with_index do |row, index|
  next if index == 0  # Пропускаем первую строку (заголовки)

  current_value = row[1]&.value.to_s.strip # 2-я колонка
  current_number = row[0]&.value.to_s.strip # 1-я колонка

  next if current_value.empty?

  description = "#{row[4]&.value.to_s.strip} #{row[5]&.value.to_s.strip}"
  characteristics = parse_characteristics(row[2]&.value.to_s.strip, row[6]&.value.to_s.strip)

  # Определяем тип компонента
  component_type = current_number[0]  # Первая буква в номере компонента

  # Если тип компонента новый (или отличается от предыдущего), вставляем строку-заголовок
  if component_type != last_category && CATEGORY_MAP.key?(component_type)
    category_name = CATEGORY_MAP[component_type]
    data << [
      "",            # Пустая ячейка
      category_name, # Название категории
      "",            # Пустая ячейка
      ""             # Пустая ячейка
    ]
  end

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
  last_category = component_type  # Обновляем текущую категорию
end

data.last[0] = format_numbers(current_numbers) unless current_numbers.empty?

# Работа с Word
FileUtils.cp(docx_path, new_docx_path)
Zip::File.open(new_docx_path) do |zip|
  document_xml = zip.find_entry("word/document.xml")

  if document_xml
    xml_content = document_xml.get_input_stream.read
    doc = Nokogiri::XML(xml_content)

    File.write("before_edit.xml", doc.to_xml)

    tables = doc.xpath("//w:tbl", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    puts "🔹 Найдено таблиц: #{tables.size}"

    tables.each do |table|
      puts "🔹 Найдено строк в таблице: #{table.xpath('.//w:tr').size}"

      data.each do |row_data|
        first_cell_value = row_data[0].to_s
        should_insert_empty_row = first_cell_value.length > 7 && first_cell_value.match?(/^(\w+)([-,])(\w+)$/)
        
        empty_row_data = nil
        if should_insert_empty_row
          first_part, separator, second_part = first_cell_value.match(/^(\w+)([-,])(\w+)$/).captures
          row_data[0] = first_part + separator
          empty_row_data = [second_part, "", row_data[2], row_data[3]]
        end
        
        new_row = Nokogiri::XML::Node.new("w:tr", doc)
        row_properties = Nokogiri::XML::Node.new("w:trPr", doc)
        row_height = Nokogiri::XML::Node.new("w:trHeight", doc)
        row_height['w:val'] = "453"
        row_height['w:hRule'] = "exact"
        row_properties.add_child(row_height)
        new_row.add_child(row_properties)
        
        formatted_cells = []
        row_data.each_with_index do |value, index|
          cell = Nokogiri::XML::Node.new("w:tc", doc)
          cell_properties = Nokogiri::XML::Node.new("w:tcPr", doc)
          paragraph = Nokogiri::XML::Node.new("w:p", doc)
          run = Nokogiri::XML::Node.new("w:r", doc)
          text_node = Nokogiri::XML::Node.new("w:t", doc)
          text_node.content = value

          run_properties = Nokogiri::XML::Node.new("w:rPr", doc)
          font = Nokogiri::XML::Node.new("w:rFonts", doc)
          font['w:ascii'] = "GOST Type A"
          font['w:hAnsi'] = "GOST Type A"
          font['w:eastAsia'] = "GOST Type A"
          font['w:cs'] = "GOST Type A"
          run_properties.add_child(font)

          size = Nokogiri::XML::Node.new("w:sz", doc)
          size['w:val'] = "28"
          run_properties.add_child(size)

          italic = Nokogiri::XML::Node.new("w:i", doc)
          run_properties.add_child(italic)

          run.add_child(run_properties)
          run.add_child(text_node)
          paragraph.add_child(run)
          cell.add_child(paragraph)
          cell.add_child(cell_properties)
          new_row.add_child(cell)
          formatted_cells << cell
        end

        table.add_child(new_row)
        puts "✅ Добавлена строка: #{row_data.inspect}"

        if should_insert_empty_row
          empty_row = Nokogiri::XML::Node.new("w:tr", doc)
          empty_row.add_child(row_properties.dup)
          empty_row_data.each_with_index do |value, index|
            cell = formatted_cells[index].dup
            cell.xpath(".//w:t").first.content = value
            empty_row.add_child(cell)
          end
          table.add_child(empty_row)
          puts "➕ Вставлена строка с перемещением значений: #{empty_row_data.inspect}"

          # Очищаем 3 и 4 колонку в строке, которая была до вставки пустой строки
          previous_row = table.xpath(".//w:tr")[table.xpath(".//w:tr").size - 2]  # Берем строку перед добавленной пустой строкой
          previous_row.xpath('.//w:tc')[2].xpath(".//w:t").first.content = ""
          previous_row.xpath('.//w:tc')[3].xpath(".//w:t").first.content = ""

          row_data[2] = ""
          row_data[3] = ""
        end
      end
    end

    File.write("after_edit.xml", doc.to_xml)
    zip.get_output_stream("word/document.xml") { |f| f.write(doc.to_xml) }
  end
end




puts "✅ Данные успешно добавлены в shablon_pr_updated.docx"
