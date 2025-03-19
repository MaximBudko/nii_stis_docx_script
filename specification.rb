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
    sheet = xlsx[0] # Берем первый лист
    data = []
    intermediate_data = []

    sheet.each_with_index do |row, index|
        next if index == 0

        current_number = row[0]&.value.to_s.strip
        current_description = row[1]&.value.to_s.strip
        current_value = row[2]&.value.to_s.strip

        next if current_description.empty?

        current_part_number = row[4]&.value.to_s.strip
        current_manufacturer = row[5]&.value.to_s.strip
        
        
        intermediate_data << {
            number: current_number,
            part_number: current_part_number,
            manufactured: current_manufacturer
        } 
    end

    intermediate = sort_by_groups(intermediate_data)

    grouped = intermediate.group_by { |entry| entry[:part_number] }
end


data = get_excel_data("test.xlsx")
pp data