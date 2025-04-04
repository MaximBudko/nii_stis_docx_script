require 'rubyXL'
require 'zip'
require 'nokogiri'
require 'fileutils'
require 'pp'
require 'stringio'


module Vedomost
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

    def self.get_excel_data(file_path)
        workbook = RubyXL::Parser.parse(file_path)
        sheet = workbook[0]
    
        data = []
        sheet.each do |row|
        row_data = row.cells.map { |cell| cell && cell.value }
        data << row_data
        end
    
        data
    end

    def self.process_data(data)
        # Оставляем только значения с индексами 0, 4 и 5
        processed_data = data.map { |row| [row[0], row[4], row[5]] }
    
        # Подсчитываем количество по 4 индексу
        count_hash = Hash.new(0)
        processed_data.each { |row| count_hash[row[1]] += 1 }
    
        # Добавляем строку с подсчетом количества
        processed_data.each do |row|
        row[3] = "#{count_hash[row[1]]}"
        end
    
        # Оставляем только уникальные строки
        unique_data = processed_data.uniq { |row| row[1] }
    
        unique_data
    end

    def self.group_data(data)
        grouped_data = []

        # Группируем данные по буквенной части нулевого индекса
        data.group_by { |row| row[0].match(/[A-Za-z]+/)[0] }.each do |key, group|
        # Проверяем наличие ключа в CATEGORY_MAP
        if CATEGORY_MAP.key?(key)
            # Получаем название группы из CATEGORY_MAP
            category_name = CATEGORY_MAP[key][group.size == 1 ? 0 : 1]
            
            # Добавляем название группы перед массивом группы
            grouped_data << ["", "", "", "", "", "", "", "", "", "", ""]
            grouped_data << ["", category_name, "", "","", "", "", "", "", "", ""]
            grouped_data.concat(group)
        else
            # Если ключ не найден, добавляем группу без названия
            grouped_data << ["", "", "", "", "", "", "", "", "", "", ""]
            grouped_data << ["", "Неизвестная категория", "", "", "", "", "", "", "", "", ""]
            grouped_data.concat(group)
        end
        end
    
        grouped_data
    end

    def self.modify_data(data)
        # Удаляем первые три массива
        data.shift(1)
    
        modified_data = []
    
        data.each do |row|
            # Создаем массив нужного размера (11 элементов)
            modified_row = Array.new(11, "")
            
            # Копируем первые два элемента
            modified_row[0] = row[0]
            modified_row[1] = row[1]
            
            # Добавляем количество в нужные позиции (6 и 9)
            modified_row[6] = row[3]
            modified_row[9] = row[3]
            
            # Добавляем производителя в последнюю позицию
            modified_row[10] = row[2]
    
            modified_data << modified_row
        end
    
        modified_data
    end

    def self.split_long_strings(data)
        modified_data = []
        
        data.each do |row|
        # Добавляем текущую строку в результат
        modified_data << row
        
        # Проверяем длину строки в индексе 10
        next unless row[10] && row[10].length > 12
        
        # Разбиваем строку на слова
        words = row[10].split
        
        # Оставляем первое слово в текущей строке
        row[10] = words.shift
        
        # Если есть оставшиеся слова, создаем новую строку
        if words.any?
            # Создаем новую строку с пустыми значениями
            new_row = Array.new(row.size, "")
            # Добавляем оставшиеся слова
            new_row[10] = words.join(" ")
            # Добавляем новую строку в результат
            modified_data << new_row
        end
        end
        
        modified_data
    end

    def self.add_iterators(data)
        modified_data = []
        page_size = 28
        first_page_size = 22
        
        # Обрабатываем первую страницу
        first_page = data.take(first_page_size)
        first_page.each_with_index do |row, index|
        row = row.dup
        row[0] = (index + 1).to_s
        modified_data << row
        end
        
        # Добавляем пустые строки, если данных меньше 22
        (first_page_size - first_page.size).times do |i|
        new_row = Array.new(11, "")
        new_row[0] = (first_page.size + i + 1).to_s
        modified_data << new_row
        end
        
        # Обрабатываем остальные страницы
        remaining_data = data[first_page_size..]
        return modified_data if remaining_data.nil?
        
        remaining_data.each_slice(page_size) do |page|
        page.each_with_index do |row, index|
            row = row.dup
            row[0] = (index + 1).to_s
            modified_data << row
        end
        
        # Добавляем пустые строки до конца страницы
        (page_size - page.size).times do |i|
            new_row = Array.new(11, "")
            new_row[0] = (page.size + i + 1).to_s
            modified_data << new_row
        end
        end
        
        modified_data
    end

    def self.group_data(data)
        # Initialize result array
        result = []
        
        # Сначала обработаем длинные строки
        processed_data = []
        current_row = nil
        
        data.each do |row|
            if current_row && row.all? { |cell| cell.to_s.empty? } && row[10].to_s.strip.length > 0
                # Это продолжение длинного названия
                current_row[10] = "#{current_row[10]} #{row[10]}"
            else
                processed_data << row
                current_row = row
            end
        end
        
        # Group data by prefix
        groups = processed_data.group_by do |row|
            next nil if row[0].nil? || row[0].empty?
            match = row[0].to_s.match(/^[A-Za-z]+/)
            match ? match[0] : 'UNKNOWN'
        end
        
        # Remove empty group if exists
        groups.delete(nil)
        
        # Sort groups
        sorted_groups = groups.sort_by do |prefix, _|
            if CATEGORY_MAP[prefix]
                [0, CATEGORY_MAP[prefix].first]
            else
                [1, prefix]
            end
        end
        
        # Process each group
        sorted_groups.each do |prefix, items|
            # Add separator
            result << Array.new(items.first.size, "")
            
            # Add group header
            header_row = Array.new(items.first.size, "")
            if CATEGORY_MAP.key?(prefix)
                category_name = CATEGORY_MAP[prefix][items.size == 1 ? 0 : 1]
            else
                category_name = "Неизвестная категория"
            end
            header_row[1] = category_name
            result << header_row
            
            # Add all items in the group
            items.each { |item| result << item }
        end
        
        result
    end

    def self.sort_within_groups(data)
        groups = data.group_by { |row| row[0].to_s.match(/^[A-Za-z]+/)[0] }
        
        sorted_groups = groups.transform_values do |group_items|
          group_items.sort_by { |item| item[1].to_s.downcase }
        end
        
        result = []
        sorted_groups.each do |prefix, items|
          result.concat(items)
        end
        
        result
    end
    
    def self.generate_docx(docx_path, xlsx_path, field_values, new_file_path)
        values = field_values
        excel_data = get_excel_data(xlsx_path)
        process_data = process_data(excel_data)
        modified_data = modify_data(process_data)
        alpha_data = sort_within_groups(modified_data)
        grouped_data = group_data(alpha_data)
        splitted_data = split_long_strings(grouped_data)
        
        data5 = add_iterators(splitted_data)

        new_docx_path = new_file_path + ".docx"
        
        begin
        sleep 0.5
        FileUtils.cp(docx_path, new_docx_path)
        rescue Errno::EACCES => e
        puts "Permission denied while copying the file: #{e.message}"
        return
        end
        retries = 1
        begin
        Zip::File.open(new_docx_path) do |zip|
            sleep(1) 
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
                        puts "Работаю :D" # Имитация работы, можно адаптировать
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
                    if index == 2
                        if value.to_s.length >= 10  # Проверяем длину текста
                          spacing = Nokogiri::XML::Node.new("w:spacing", doc)
                          spacing['w:val'] = "-10"  # Значение уплотнения
                          run_properties.add_child(spacing)
                        end
                    end
                        
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
        rescue Errno::EACCES => e
        if retries > 0
            retries -= 1
            sleep(1)
            retry
        else
            puts "Permission denied while processing the file: #{e.message}"
        end
        end
    end
end