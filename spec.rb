require 'rubyXL'
require 'zip'
require 'nokogiri'
require 'fileutils'
require 'pp'
require 'stringio'


module Spec
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


  def self.sort_by_groups(array)
      grouped = array.group_by { |item| item[:number][/^[A-Za-z]+/] }
      
      # Сортируем только по числовой части внутри каждой группы
      sorted = grouped.transform_values do |group|
        group.sort_by { |item| item[:number][/\d+/].to_i }
      end
    
      # Восстанавливаем порядок групп из исходного массива
      array.map { |item| sorted[item[:number][/^[A-Za-z]+/]].shift }
  end


  def self.get_excel_data(excel_path)
      xlsx = RubyXL::Parser.parse(excel_path)
      sheet = xlsx[0]
      intermediate_data = []

      # Сначала собираем все данные и группируем по part_number
      part_number_groups = {}
      
      sheet.each_with_index do |row, index|
          next if index == 0

          current_number = row[0]&.value.to_s.strip
          current_description = row[1]&.value.to_s.strip
          current_value = row[2]&.value.to_s.strip
          current_part_number = row[4]&.value.to_s.strip
          current_manufacturer = row[5]&.value.to_s.strip

          next if current_description.empty?
          
          # Группируем по part_number
          part_number_groups[current_part_number] ||= {
              numbers: [],
              manufactured: current_manufacturer,
              part_number: current_part_number
          }
          
          part_number_groups[current_part_number][:numbers] << current_number
      end

      # Преобразуем группы в промежуточные данные
      part_number_groups.each do |part_number, group|
          intermediate_data << {
              number: group[:numbers].join(", "),  # Сохраняем все номера
              part_number: group[:part_number],
              manufactured: group[:manufactured],
              count: group[:numbers].length  # Добавляем количество
          }
      end

      # Сортировка и группировка
      grouped = intermediate_data.group_by { |entry| entry[:part_number] }
      
      # Сортировка по префиксу
      sorted_by_prefix = grouped.sort_by do |_, values|
          values.first[:number].to_s[/^[A-Za-z]+/]
      end.to_h

      # Дальнейшая сортировка по первой букве part_number
      final_sorted = {}
      current_prefix = nil
      current_group = {}

      sorted_by_prefix.each do |part_number, values|
          prefix = values.first[:number].to_s[/^[A-Za-z]+/]
          if current_prefix != prefix && !current_group.empty?
              final_sorted.merge!(sort_group_by_part_number(current_group))
              current_group = {}
          end
          current_prefix = prefix
          current_group[part_number] = values
      end

      final_sorted.merge!(sort_group_by_part_number(current_group)) unless current_group.empty?
      final_sorted
  end

  def self.format_sequence_numbers(numbers_string)
    numbers = numbers_string.split(/,\s*/)
    prefix = numbers.first[/^[A-Za-z]+/]
    
    # Сортируем номера по числовой части
    sorted_numbers = numbers.sort_by { |num| num[/\d+/].to_i }
    
    # Находим последовательности
    sequences = []
    current_seq = [sorted_numbers.first]
    
    sorted_numbers[1..-1].each do |num|
      current_num = num[/\d+/].to_i
      prev_num = current_seq.last[/\d+/].to_i
      
      if current_num == prev_num + 1
        current_seq << num
      else
        sequences << format_single_sequence(current_seq)
        current_seq = [num]
      end
    end
    sequences << format_single_sequence(current_seq)
    
    sequences.join(", ")
  end

  def self.format_single_sequence(sequence)
    return sequence.first if sequence.length == 1
    return sequence.join(", ") if sequence.length == 2
    "#{sequence.first}-#{sequence.last}"
  end

  def self.split_long_numbers(row, max_length = 10)
    numbers = row[6].split(/,\s*/)
    result = []
    current_line = []
    current_length = 0

    numbers.each_with_index do |num, index|
      if (current_length + num.length + 2) <= max_length
        current_line << num
        current_length = current_line.join(", ").length
      else
        # Добавляем запятую только если это не последняя строка
        result << current_line.join(", ") + (index < numbers.length - 1 ? "," : "")
        current_line = [num]
        current_length = num.length
      end
    end

    # Добавляем последнюю строку без запятой
    result << current_line.join(", ") unless current_line.empty?

    # Форматируем результат
    first_row = row.clone
    first_row[6] = result.first

    additional_rows = result[1..-1].map do |numbers|
      ["", "", "", "", "", "", numbers]
    end

    # Добавляем пустую строку только после полного блока
    [first_row] + additional_rows + [["", "", "", "", "", "", ""]]
  end

  def self.format_to_array(hash, start_iter = 1)
    result = []
    current_prefix = nil
    current_manufacturer = nil
    is_first_item = true
    iter = start_iter
    
    hash.each do |part_number, values|
      prefix = values.first[:number][/^[A-Za-z]+/]
      manufactured = values.first[:manufactured]
      
      # Получаем количество из списка номеров
      numbers_array = values.first[:number].split(/,\s*/)
      quantity = numbers_array.length
      
      # Форматируем последовательности номеров
      formatted_numbers = format_sequence_numbers(values.first[:number])
      
      # Сброс флага при смене производителя
      if current_manufacturer != manufactured
        is_first_item = true
        current_manufacturer = manufactured
      end

      if quantity > 1
        # Множественные элементы
        if is_first_item
          result << ["", "", "", "", "#{CATEGORY_MAP[prefix]&.[](1)} #{manufactured}", "", ""]
        end
        base_row = ["", "", iter.to_s, "", part_number, quantity.to_s, formatted_numbers]
        result.concat(split_long_numbers(base_row))
      else
        # Одиночные элементы
        if is_first_item
          result << ["", "", iter.to_s, "", "#{CATEGORY_MAP[prefix]&.[](0)} #{part_number}", quantity.to_s, formatted_numbers]
          result << ["", "", "", "", manufactured, "", ""]
        else
          result << ["", "", iter.to_s, "", part_number, quantity.to_s, formatted_numbers]
        end
        result << ["", "", "", "", "", "", ""]
      end
      
      is_first_item = false
      iter += 1
    end

    fix_comma_arrays(result)
  end

  def self.fix_comma_arrays(array)
    result = []
    i = 0
    
    while i < array.length
      current_row = array[i]
      
      # Проверяем, является ли текущая строка строкой с запятой
      if current_row[6] == ","
        # Получаем следующую строку, если она существует
        next_row = array[i + 1]
        
        if next_row
          # Перемещаем значения из текущей строки в следующую
          next_row[2] = current_row[2] if current_row[2] != ""    # Номер позиции
          next_row[4] = current_row[4] if current_row[4] != ""    # Парт номер
          next_row[5] = current_row[5] if current_row[5] != ""    # Количество
          
          # Добавляем следующую строку с перенесенными значениями в результат
          result << next_row
          # Пропускаем обработку следующей строки, так как мы её уже добавили
          i += 2
          next
        else
          # Если следующей строки нет, просто пропускаем текущую
          i += 1
          next
        end
      else
        # Добавляем текущую строку в результат
        result << current_row
        i += 1
      end
    end
    
    result
  end

  private

  def self.sort_group_by_part_number(group)
    group.sort_by do |part_number, _|
      # Если начинается с цифры, добавляем 'z' впереди для корректной сортировки
      first_char = part_number[0]
      first_char =~ /\d/ ? "z#{part_number}" : part_number
    end.to_h
  end

  def self.generate_spec(docx_path, xlsx_path, field_values, new_file_path, input_int)
    xlsx = get_excel_data(xlsx_path)
    values = field_values

    data5 = format_to_array(xlsx, input_int)

    new_docx_path = new_file_path + ".docx"

    begin
      sleep 0.5
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
        sleep(1)
        if document_xml
          xml_content = document_xml.get_input_stream.read
          doc = Nokogiri::XML(xml_content)
          tables = doc.xpath("//w:tbl", "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

          table = tables[0]
            
              start_time = Time.now
              while Time.now - start_time < 5
                # Ваше действие здесь
                puts "Действие выполняется..."# Для имитации длительного выполнения действия, можно адаптировать под нужды
              end
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
          
        zip.get_output_stream("word/document.xml") { |f| f.write(doc.to_xml) }
        end
      end
  end
end