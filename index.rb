require 'roo'
require 'caracal'
require 'securerandom'

# Словарь замен
UNIT_MAP = {
  'n' => 'нФ',
  'u' => 'мкФ',
  'm' => 'МОм',
  'k' => 'кОм',
  'p' => 'пФ'
}

def excel_to_docx(excel_file)
  unless File.exist?(excel_file)
    puts "Ошибка: файл #{excel_file} не найден!"
    return
  end

  xlsx = Roo::Excelx.new(excel_file)
  docx_file = "output_#{SecureRandom.hex(4)}.docx"

  Caracal::Document.save(docx_file) do |docx|
    headers = ["Поз. обозна-чение","Наименование","Кол.","Примечание"] # Заголовки
    data = []
    last_value = nil
    count = 1
    current_numbers = []

    xlsx.each_row_streaming(offset: 1) do |row|
      next if row.empty?

      current_value = row[1]&.value.to_s.strip # 2-я колонка
      current_number = row[0]&.value.to_s.strip # 1-я колонка

      if current_value.empty?
        next # Пропуск пустых строк
      end

      # Объединяем 5-й и 6-й столбцы в "Описание"
      description = "#{row[4]&.value.to_s.strip} #{row[5]&.value.to_s.strip}"

      # Создаем "Характеристики"
      characteristics = parse_characteristics(row[2]&.value.to_s.strip, row[6]&.value.to_s.strip)

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

    docx.table([headers] + data)
  end

  puts "Файл #{docx_file} успешно создан!"
end

# Функция форматирования номеров
def format_numbers(numbers)
  return numbers.first if numbers.size == 1

  sorted_numbers = numbers.sort_by { |num| num[/\d+/].to_i rescue num } # Сортировка по числам
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

# Функция обработки "Характеристик"
def parse_characteristics(value, tolerance)
  return "" if value.empty?

  value = value.gsub(",", ".") # Заменяем запятую на точку для корректного парсинга

  unit = value[/[a-zA-Z]+/] # Извлекаем единицу измерения
  number = value[/\d+(\.\d+)?/] # Извлекаем число (с десятичной точкой)
  dnp = value.include?("DNP") ? " DNP" : ""

  unit = UNIT_MAP[unit] || unit # Подставляем русское обозначение

  formatted_value = number ? "#{number} #{unit}" : value
  formatted_value += "±#{tolerance}" unless tolerance.empty?

  # Заменяем точку на запятую в числе
  formatted_value.gsub(".", ",") + dnp
end

# Запуск
excel_to_docx('Test.xlsx')