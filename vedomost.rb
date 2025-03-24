require 'rubyXL'
require 'zip'
require 'nokogiri'
require 'fileutils'
require 'pp'
require 'stringio'

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
        grouped_data << ["", "", "", ""]
        grouped_data << ["", category_name, "", ""]
        grouped_data.concat(group)
      else
        # Если ключ не найден, добавляем группу без названия
        grouped_data << ["", "", "", ""]
        grouped_data << ["", "Неизвестная категория", "", ""]
        grouped_data.concat(group)
      end
    end
  
    grouped_data
end

def self.modify_data(data)
    # Удаляем первые три массива
    data.shift(3)
  
    modified_data = []
  
    data.each do |row|
      # Добавляем пустые строки с 2 по 5 индекс
      modified_row = row[0..1] + ["", "", "", ""]
  
      # Добавляем строку с количеством
      modified_row << row[3]
  
      # Добавляем две пустые строки
      modified_row << ""
      modified_row << ""
  
      # Добавляем еще одну строку с количеством
      modified_row << row[3]
  
      # Добавляем строку производителя, которая была под индексом 2
      modified_row << row[2]
  
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

file_path = 'test.xlsx'
excel_data = get_excel_data(file_path)
process_data = process_data(excel_data)
grouped_data = group_data(process_data)
modified_data = modify_data(grouped_data)
splitted_data = split_long_strings(modified_data)
final_data = add_iterators(splitted_data)

pp final_data