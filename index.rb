require 'rubyXL'
require 'zip'
require 'nokogiri'
require 'fileutils'
require 'pp'
require 'stringio'

module ExcelToDocx
  # Словарь замен единиц измерения
  UNIT_CATEGORY = {
    "p" => "п",
    "n" => "н",
    "u" => "мк",
    "m" => "м",
    "k" => "к",
    "M" => "М"
  }

  # Словарь соответствий типов компонентов
  CATEGORY_MAP = {
    "C" => 	["Конденсатор", "Конденсаторы"],
    "D" => 	["Микросхема", "Микросхемы"],
    "DA" => ["Микросхема аналоговая",	"Микросхемы аналоговые"],
    "E" =>	["Элемент", "Элементы"],
    "F" =>	["Предохранитель", "Предохранители"],
    "G" => 	["Генератор", "Генераторы"],
    "GB" =>	["Батарея литиевая", "Батареи литиевые"],
    "H" =>	["Индикатор", "Индикаторы"],
    "X" => 	["Соединитель", "Соединители"],
    "K" =>	["Реле", "Реле"],
    "L" => 	["Дроссель", "Дроссели"],
    "R" => 	["Резистор", "Резисторы"],
    "S" =>	["Кнопка тактовая", "Кнопки тактовые"],
    "T" => 	["Трансформатор", "Трансформаторы"],
    "U" => 	["Модуль", "Модули"],
    "VD" =>	["Диод", "Диоды"],
    "VT" =>	["Транзистор", "Транзисторы"],
    "P" =>	["Реле", "Реле"],
    "FA" =>	["Предохранитель", "Предохранители"],
    "Z" =>	["Кварцевый резонатор", "Кварцевые резонаторы"]
  }

  ALIAS = {
    "J" => "X",
    "HL" => "H",
    "SB" => "S"
  }

  def self.parse_characteristics(value, tolerance, current_number)
    regxp_current_number = current_number[/\A[a-zA-Z]+/]
    value = value.gsub(",", ".") # Заменяем запятую на точку
    unit_second = ""
    if regxp_current_number == "C"
      unit_second = "Ф"
    elsif regxp_current_number == "D" 
      unit_second = ""
    elsif regxp_current_number == "DA"
      unit_second = ""
    elsif regxp_current_number == "E"
      unit_second = ""
    elsif regxp_current_number == "F"
      unit_second = ""
    elsif regxp_current_number == "G"
      unit_second = ""
    elsif regxp_current_number == "GB"
      unit_second = ""
    elsif regxp_current_number == "H"
      unit_second = ""
    elsif regxp_current_number == "X"
      unit_second = ""
    elsif regxp_current_number == "K"
      unit_second = ""
    elsif regxp_current_number == "L"
      unit_second = "Гн"
    elsif regxp_current_number == "R"
      unit_second = "Ом"
    elsif regxp_current_number == "S"
      unit_second = ""
    elsif regxp_current_number == "T"
      unit_second = ""
    elsif regxp_current_number == "U"
      unit_second = ""
    elsif regxp_current_number == "VD"
      unit_second = ""
    elsif regxp_current_number == "VT"
      unit_second = ""
    elsif regxp_current_number == "P"
      unit_second = ""
    elsif regxp_current_number == "FA"
      unit_second = ""
    elsif regxp_current_number == "Z"
      unit_second = ""
    end

    unit_first = value[/[a-zA-Z]+/] # Извлекаем единицу измерения
    number = value[/\d+(\.\d+)?/] # Извлекаем число
    dnp = value.include?("DNP") ? " DNP" : ""

    unit = UNIT_CATEGORY[unit_first] || ""
    if regxp_current_number == "DA" || regxp_current_number == "F" || regxp_current_number == "D" ||
      regxp_current_number == "G" || regxp_current_number == "K" || regxp_current_number == "HL" ||
      regxp_current_number == "SB" || regxp_current_number == "U" || regxp_current_number == "VD"||
      regxp_current_number == "VT" || regxp_current_number == "X"

      return "#{value.include?("DNP") ? " DNP" : ""}"
    end
    if UNIT_CATEGORY.include?(unit_first) == false
      if value.match?(/^[\d[:punct:]\s]+$/) == true || value.match?(/^[\d[:punct:]\s]+_DNP$/) == true 
        formatted_value = "#{number == " " ? "0" : number } #{unit}#{unit_second}"
        formatted_value += "±#{tolerance}" unless tolerance.nil? || tolerance.empty?
        return formatted_value.gsub(".", ",") + dnp
      else
        return "#{value.include?("DNP") ? " DNP" : ""}"
      end
    end

    formatted_value = "#{number == " " ? "0" : number } #{unit}#{unit_second}"
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
      if row[0].scan(/\S/).length > 7
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
    grouped_data = {}
    category_replacements = {}
  
    # 1️⃣ Формируем категории в порядке поступления
    data.each do |row|
      category_key = row[0].to_s[/\A[a-zA-Z]+/] || ''
  
      # Проверяем, есть ли замена в ALIAS
      if ALIAS.key?(category_key)
        new_key = ALIAS[category_key]
  
        # Если новая категория уже существует, переносим туда
        if grouped_data.key?(new_key)
          grouped_data[new_key] += grouped_data.delete(category_key) if grouped_data.key?(category_key)
        end
  
        category_replacements[category_key] = new_key
        category_key = new_key
      end
  
      category_name = CATEGORY_MAP[category_key] || category_key
      grouped_data[category_name] ||= [] # Создаем категорию, если её нет
      grouped_data[category_name] << row
    end
  
    # 2️⃣ Сохраняем порядок следования категорий
    ordered_categories = data.map { |row| row[0].to_s[/\A[a-zA-Z]+/] || '' }
                             .uniq
                             .map { |key| CATEGORY_MAP[key] || key }
  
    result = []
    ordered_categories.each do |category|
      next unless grouped_data.key?(category) # Пропускаем пустые категории
  
      items = grouped_data[category]
      next if items.empty? # Пропускаем заголовки без элементов
  
      selected_key = category.length > 1 ? category[1] : category[0]
  
      result << ["", "", "", ""] unless result.empty? || result.last == ["", "", "", ""]
      result << ["", selected_key, "", ""]
      result.concat(items)
    end
  
    # 🔥 Убираем лишний заголовок в конце (если есть)
    result.pop while result.any? && result.last == ["", "", "", ""]
  
    result
  end
  
  def self.move_first_to_end(arr)
    empty_row = ["", "", "", ""]
    arr.insert(0, empty_row.dup)
    arr.push(arr.shift)
  end

  def self.insert_empty_and_move(data)
    empty_row = ["", "", "", ""]
    index = 24 # так как индексация с 0, 24-й элемент имеет индекс 23
    index_for_move = 23

    while index < data.length

      if data[index] == empty_row

      elsif data[index - 1][1] != "" && data[index - 1][2] == "" && data[index - 1][3] == ""
        data.insert(index - 1, empty_row.dup)
        data.insert(index, empty_row.dup)
      else
        data.insert(index, empty_row.dup) # вставляем копию пустого массива
      end
      index += 30 # сдвигаем индекс на 29 позиций (учиxтывая вставленный элемент)
    end
    
    data
  end

  def self.sort_by_groups(array)
    grouped = array.group_by { |item| item[:number][/^[A-Za-z]+/] }
    
    # Сортируем только по числовой части внутри каждой группы
    sorted = grouped.transform_values do |group|
      group.sort_by { |item| item[:number][/\d+/].to_i }
    end
  
    # Восстанавливаем порядок групп из исходного массива
    array.map { |item| sorted[item[:number][/^[A-Za-z]+/]].shift }
  end

  def self.generate_docx(docx_path, xlsx_path, field_values, new_file_path)
    xlsx = RubyXL::Parser.parse(xlsx_path)
    sheet = xlsx[0] # Берем первый лист
    values = field_values
    data = []
    last_value = nil
    last_qnt = nil
    count = 1
    current_numbers = []
    last_category = nil 
    intermediate_data = []

    sheet.each_with_index do |row, index|
      next if index == 0  # Пропускаем первую строку (заголовки)

      current_value = row[1]&.value.to_s.strip # 2-я колонка
      current_number = row[0]&.value.to_s.strip # 1-я колонка
      current_qnt = row[2]&.value.to_s.strip # 3-я колонка
      next if current_value.empty?

      description = "#{row[4]&.value.to_s.strip} #{row[5]&.value.to_s.strip}"
      characteristics = parse_characteristics(row[2]&.value.to_s.strip, row[6]&.value.to_s.strip, current_number)
      intermediate_data << {
        number: current_number,
        value: current_value,
        qnt: current_qnt,
        description: description,
        characteristics: characteristics
      }
    end
    intermediate = sort_by_groups(intermediate_data)
    intermediate.each do |entry|
      if entry[:value] == last_value && entry[:qnt] == last_qnt
        count += 1
        current_numbers << entry[:number]
        data.last[2] = count.to_s
      else
        data.last[0] = format_numbers(current_numbers) unless current_numbers.empty?
        count = 1
        current_numbers = [entry[:number]]
        data << [
          entry[:number],
          entry[:description],
          "1",  # Количество
          entry[:characteristics]
        ]
      end
      last_qnt = entry[:qnt]
      last_value = entry[:value]
    end

    data.last[0] = format_numbers(current_numbers) unless current_numbers.empty?
    data1 = group_by_category(data)
    data2 = process_array(data1)
    data3 = move_first_to_end(data2)
    data5 = insert_empty_and_move(data3)
    new_docx_path = new_file_path + ".docx"
    
    begin
      FileUtils.cp(docx_path, new_docx_path)
    rescue Errno::EACCES => e
      puts "Permission denied while copying the file: #{e.message}"
      return
    end

    Zip::File.open(new_docx_path) do |zip|

      zip.glob('word/{header,footer}*.xml').each do |entry|
        xml_content = entry.get_input_stream.read
        doc = Nokogiri::XML(xml_content)
        namespaces = { 'w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' }

        doc.xpath('//w:t', namespaces).each do |node|
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
              

              borders = Nokogiri::XML::Node.new("w:tcBorders", doc)
              ["w:top", "w:bottom", "w:left", "w:right"].each do |side|
                border = Nokogiri::XML::Node.new(side, doc)
                border['w:val'] = "single"  # Сплошная линия
                border['w:sz'] = "10"        # Толщина границы
                border['w:color'] = "000000" # Черный цвет
                borders.add_child(border)
              end
              cell_properties.add_child(borders)

              v_align = Nokogiri::XML::Node.new("w:vAlign", doc)
              v_align['w:val'] = 'center'  
              cell_properties.add_child(v_align)

              paragraph = Nokogiri::XML::Node.new("w:p", doc)
              paragraph_properties = Nokogiri::XML::Node.new("w:pPr", doc)
              paragraph.add_child(paragraph_properties)
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
  rescue Errno::EACCES => e
    puts "Permission denied while processing the file: #{e.message}"
  end

end

