require 'rubyXL'
require 'zip'
require 'nokogiri'
require 'fileutils'
require 'pp'
require 'stringio'


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


def sort_by_groups(array)
    grouped = array.group_by { |item| item[:number][/^[A-Za-z]+/] }
    
    # Сортируем только по числовой части внутри каждой группы
    sorted = grouped.transform_values do |group|
      group.sort_by { |item| item[:number][/\d+/].to_i }
    end
  
    # Восстанавливаем порядок групп из исходного массива
    array.map { |item| sorted[item[:number][/^[A-Za-z]+/]].shift }
end


def get_excel_data(excel_path)
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

def format_to_array(hash, start_iter = 1)
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
      result << ["", "", iter.to_s, "", part_number, quantity.to_s, values.first[:number]]
      result << ["", "", "", "", "", "", ""]
    else
      # Одиночные элементы
      if is_first_item
        result << ["", "", iter.to_s, "", "#{CATEGORY_MAP[prefix]&.[](0)} #{part_number}", quantity.to_s, values.first[:number]]
        result << ["", "", "", "", manufactured, "", ""]
      else
        result << ["", "", iter.to_s, "", part_number, quantity.to_s, values.first[:number]]
      end
      result << ["", "", "", "", "", "", ""]
    end
    
    is_first_item = false
    iter += 1
  end

  result
end

private

def sort_group_by_part_number(group)
  group.sort_by do |part_number, _|
    # Если начинается с цифры, добавляем 'z' впереди для корректной сортировки
    first_char = part_number[0]
    first_char =~ /\d/ ? "z#{part_number}" : part_number
  end.to_h
end


data = get_excel_data("test.xlsx")
formatted_data = format_to_array(data, 1) # Start numbering from 1
pp formatted_data
#pp data