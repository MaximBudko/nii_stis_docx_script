require 'roo'
require 'write_xlsx'

module CompressedBom
    def self.compress_numbers(numbers)
        return "" if numbers.empty?

        sorted_numbers = numbers.compact.sort_by { |n| n[/\d+/]&.to_i || 0 }
        ranges = []
        range_start = range_end = sorted_numbers.first
        
        sorted_numbers.each_cons(2) do |a, b|
            if b[/\d+/]&.to_i == a[/\d+/]&.to_i + 1
                range_end = b
            else
                ranges << (range_start == range_end ? range_start : "#{range_start}-#{range_end}")
                range_start = range_end = b
            end
        end
        ranges << (range_start == range_end ? range_start : "#{range_start}-#{range_end}")
        ranges.join(', ')
    end

    def self.process_excel(input_file, output_file="output.xlsx")
        xlsx = Roo::Excelx.new(input_file)
        sheet = xlsx.sheet(0)

        # Читаем заголовки
        headers = sheet.row(1) rescue []
        
        # Хеш для агрегации данных по 5-й колонке
        aggregated_data = {}

        sheet.each_row_streaming(offset: 1, pad_cells: true) do |row|
            key = row[4]&.value  # Уникальное значение из пятой колонки
            next if key.nil? # Пропускаем пустые строки

            number = row[0]&.value.to_s # Номер из первой колонки
            quantity = row[3]&.value.to_i # Количество из четвертой
            
            if aggregated_data.key?(key)
                aggregated_data[key][:quantity] += 1
                aggregated_data[key][:numbers] << number
            else
                aggregated_data[key] = {
                    quantity: 1,
                    numbers: [number],
                    full_row: row.map { |cell| cell&.value }
                }
            end
        end

        # Подготавливаем данные для записи
        result_data = aggregated_data.map do |key, value|
            new_row = value[:full_row].dup  # Дублируем, чтобы не менять исходные данные
            new_row[0] = compress_numbers(value[:numbers])
            new_row[3] = value[:quantity]
            new_row
        end

        # Создаём новый Excel-файл
        workbook = WriteXLSX.new(output_file)
        worksheet = workbook.add_worksheet

        # Записываем заголовки (если есть)
        worksheet.write_row(0, 0, headers) unless headers.empty?

        # Записываем данные
        result_data.each_with_index do |row, row_index|
            worksheet.write_row(row_index + 1, 0, row)
        end
        workbook.close
    end
end
