require 'rubyXL'
require 'zip'
require 'nokogiri'
require 'fileutils'
require 'pp'

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
  'L' => 'Дросcли',
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

  # Определяем тип компонента по первым символам номера
  component_type = CATEGORY_MAP.keys.find { |key| current_number.start_with?(key) }  # Поиск совпадения по начальной части строки

  # Если значение то же, добавляем к текущим данным
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

def insert_empty_rows(data)
  empty_row = ["", "", "", ""]
  index = 24 # так как индексация с 0, 24-й элемент имеет индекс 23
  
  while index < data.length
    data.insert(index, empty_row.dup) # вставляем копию пустого массива
    index += 30 # сдвигаем индекс на 29 позиций (учитывая вставленный элемент)
  end
  
  data
end


#data.each { |sub_array| puts sub_array.inspect }

def process_array(data)
  processed_data = []

  data.each do |row|
    if row[0].length > 7
      parts = row[0].rpartition(/[-,]/) # Разделяем по последнему '-' или ','
      if parts[1] != ""
        processed_data << [parts[0] + parts[1], row[1], "", ""]
        processed_data << [parts[2], "", row[2], row[3]]
      else
        processed_data << row # Если не удалось разделить, оставляем как есть
      end
    else
      processed_data << row
    end
  end

  processed_data
end

def group_by_category(data)
  grouped_data = Hash.new { |hash, key| hash[key] = [] }

  # Группируем элементы по первой букве первого элемента (категория)
  data.each do |row|
    category_key = row[0][0..1] # Берем первые 1-2 символа (например, "R", "C", "DA")
    category_key = CATEGORY_MAP.keys.include?(category_key) ? category_key : category_key[0] # Проверяем, если нет двухбуквенного кода, берем первую букву
    category_name = CATEGORY_MAP[category_key] || 'Неизвестная категория'
    grouped_data[category_name] << row
  end

  # Формируем новый массив с заголовками
  result = []
  grouped_data.each do |category, items|
    result << ["", "", "", ""]
    result << ["", category, "", ""]
    result.concat(items)
  end

  result
end

def move_first_to_end(arr)
  arr.push(arr.shift)
end

data1 = group_by_category(data)
data2 = process_array(data1)
data3 = move_first_to_end(data2)
data4 = insert_empty_rows(data3)





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

      data4.each do |row_data|  
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
      end
    end

    File.write("after_edit.xml", doc.to_xml)
    zip.get_output_stream("word/document.xml") { |f| f.write(doc.to_xml) }
  end
end




puts "✅ Данные успешно добавлены в shablon_pr_updated.docx"