require 'rubyXL'
require 'zip'
require 'nokogiri'
require 'fileutils'
require 'pp'
require 'stringio'

module ExcelToDocx
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
    'L' => 'Дроссели',
    'R' => 'Резисторы',
    'SB' => 'Кнопки тактовые',
    'U' => 'Модули',
    'VD' => 'Диоды',
    'VT' => 'Транзисторы',
    'X' => 'Соединители'
  }

  def self.parse_characteristics(value, tolerance)
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

  def self.format_numbers(numbers)
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
        ranges << (temp_range.size > 2 ? "#{temp_range.first}-#{temp_range.last}" : temp_range.join(", "))
        temp_range = [curr]
      end
    end

    ranges << (temp_range.size > 2 ? "#{temp_range.first}-#{temp_range.last}" : temp_range.join(", ")) unless temp_range.empty?
    ranges.join(", ")
  end

  def self.process_array(data)
    processed_data = []

    data.each do |row|
      if row[0].length > 8
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

  def self.group_by_category(data)
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

  def self.move_first_to_end(arr)
    arr.push(arr.shift)
  end

  def self.insert_empty_and_move(data)
    empty_row = ["", "", "", ""]
    index = 24 # так как индексация с 0, 24-й элемент имеет индекс 23
    index_for_move = 23

    while index < data.length
      if data[index - 1][1] != "" && data[index - 1][2] == "" && data[index - 1][3] == ""
        data.insert(index - 1, empty_row.dup)
        data.insert(index, empty_row.dup)
      else
        data.insert(index, empty_row.dup) # вставляем копию пустого массива
      end
      index += 30 # сдвигаем индекс на 29 позиций (учиxтывая вставленный элемент)
    end
    
    data
  end

  def self.generate_docx(docx_path, xlsx_path, field_values, new_file_path)

    xlsx = RubyXL::Parser.parse(xlsx_path)
    sheet = xlsx[0] # Берем первый лист

    values = field_values
    data = []
    last_value = nil
    count = 1
    current_numbers = []
    last_category = nil 

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

    data1 = group_by_category(data)
    data2 = process_array(data1)
    data3 = move_first_to_end(data2)
    data5 = insert_empty_and_move(data3)
    new_docx_path = new_file_path + ".docx"
    
    FileUtils.cp(docx_path, new_docx_path)
    Zip::File.open(new_docx_path) do |zip|

      zip.glob('word/{header,footer}*.xml').each do |entry|
        xml_content = entry.get_input_stream.read
        doc = Nokogiri::XML(xml_content)
        namespaces = { 'w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' }

        doc.xpath('//w:t', namespaces).each do |node|
          puts node
          text = node.text.strip
          if text.empty?
            node.content = ""
            parent_run = node.ancestors('w:r').first
          else
            values.each do |key, value|
              if text.include?(key)
                node.content = value
              end
            end
          end
        end
        zip.get_output_stream(entry.name) { |f| f.write(doc.to_xml) }
      end

      document_xml = zip.find_entry("word/document.xml")

      if document_xml
        xml_content = document_xml.get_input_stream.read
        doc = Nokogiri::XML(xml_content)
        tables = doc.xpath("//w:tbl", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

        tables.each do |table|

          data5.each do |row_data|  
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
          end
        end
       zip.get_output_stream("word/document.xml") { |f| f.write(doc.to_xml) }
      end
    end
  end
end

