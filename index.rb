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
  'C' => 'Конденсаторы',
  'R' => 'Резисторы',
  'L' => 'Катушки индуктивности',
  'D' => 'Диоды',
  'Q' => 'Транзисторы',
  'U' => 'Микросхемы'
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
      
        # Устанавливаем высоту строки 0,8 см
        row_properties = Nokogiri::XML::Node.new("w:trPr", doc)
        row_height = Nokogiri::XML::Node.new("w:trHeight", doc)
        row_height['w:val'] = "453"  # 0.8 см (800 twips)
        row_height['w:hRule'] = "exact" # Фиксированная высота
        row_properties.add_child(row_height)
        new_row.add_child(row_properties)
      
        row_data.each_with_index do |value, index|
          cell = Nokogiri::XML::Node.new("w:tc", doc) # Создаем ячейку
          cell_properties = Nokogiri::XML::Node.new("w:tcPr", doc) # Свойства ячейки
      
          # Устанавливаем границы
          borders = Nokogiri::XML::Node.new("w:tcBorders", doc)
      
          # Верхняя и нижняя граница для всех ячеек
          top_border = Nokogiri::XML::Node.new("w:top", doc)
          top_border['w:val'] = "single"
          top_border['w:space'] = "0"
          top_border['w:size'] = "4"  # Толщина линии
          top_border['w:space'] = "0"
      
          bottom_border = Nokogiri::XML::Node.new("w:bottom", doc)
          bottom_border['w:val'] = "single"
          bottom_border['w:space'] = "0"
          bottom_border['w:size'] = "4"
      
          borders.add_child(top_border)
          borders.add_child(bottom_border)
      
          # Для первой и четвертой колонки добавляем левую и правую границу
          if index == 0 || index == 3
            left_border = Nokogiri::XML::Node.new("w:left", doc)
            left_border['w:val'] = "single"
            left_border['w:space'] = "0"
            left_border['w:size'] = "4"
      
            right_border = Nokogiri::XML::Node.new("w:right", doc)
            right_border['w:val'] = "single"
            right_border['w:space'] = "0"
            right_border['w:size'] = "4"
      
            borders.add_child(left_border)
            borders.add_child(right_border)
          end
      
          # Добавляем границы к ячейке
          cell_properties.add_child(borders)
          cell.add_child(cell_properties)
      
          # Добавляем содержимое ячейки
          paragraph = Nokogiri::XML::Node.new("w:p", doc) # Создаем параграф
          run = Nokogiri::XML::Node.new("w:r", doc) # Создаем run (контейнер для текста)
          text_node = Nokogiri::XML::Node.new("w:t", doc) # Создаем текстовый узел
      
          text_node.content = value.empty? ? "" : value
      
          # Применение стиля шрифта GOST type A, размер 14, курсив
          run_properties = Nokogiri::XML::Node.new("w:rPr", doc)
          font = Nokogiri::XML::Node.new("w:rFonts", doc)
          font['w:ascii'] = "GOST Type A"
          font['w:hAnsi'] = "GOST Type A"
          font['w:eastAsia'] = "GOST Type A"
          font['w:cs'] = "GOST Type A"
          run_properties.add_child(font)
      
          size = Nokogiri::XML::Node.new("w:sz", doc)
          size['w:val'] = "28"  # Устанавливаем размер шрифта 14 (в половинных пунктах)
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
