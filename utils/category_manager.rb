require 'json'
require 'fileutils'

module CategoryManager
    DICTIONARY_PATH = File.expand_path('../resource/dictionaries.json', __dir__)
  
    class << self
      def load_categories
        ensure_dictionary_exists
        JSON.parse(File.read(DICTIONARY_PATH))['categories']
      rescue JSON::ParserError => e
        puts "Ошибка при чтении файла словаря: #{e.message}"
        {}
      end
  
      def save_categories(categories)
        FileUtils.mkdir_p(File.dirname(DICTIONARY_PATH))
        File.write(DICTIONARY_PATH, JSON.pretty_generate({ 'categories' => categories }))
      end
  
      def add_category(key, singular, plural)
        categories = load_categories
        return false if categories.key?(key)
        
        categories[key] = [singular, plural]
        save_categories(categories)
        true
      end
  
      def remove_category(key)
        categories = load_categories
        return false unless categories.key?(key)
        
        categories.delete(key)
        save_categories(categories)
        true
      end
  
      def update_category(key, singular, plural)
        categories = load_categories
        return false unless categories.key?(key)
        
        categories[key] = [singular, plural]
        save_categories(categories)
        true
      end
  
      private
  
      def ensure_dictionary_exists
        unless File.exist?(DICTIONARY_PATH)
          default_categories = {
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
          save_categories(default_categories)
        end
      end
    end
  end