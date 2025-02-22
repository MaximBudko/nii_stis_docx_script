require 'gtk3'

# Создаем новое окно
window = Gtk::Window.new

# Назначаем размеры окна
window.set_default_size(300, 200)
window.set_title("Пример приложения на GTK3")

# Создаем кнопку
button = Gtk::Button.new(label: "Нажми меня")

# Устанавливаем обработчик события для кнопки
button.signal_connect("clicked") do
  puts "Кнопка нажата!"
end

# Добавляем кнопку в окно
window.add(button)

# Настроим закрытие окна
window.signal_connect("destroy") do
  Gtk.main_quit
end

# Отображаем все виджеты
window.show_all

# Запуск главного цикла приложения
Gtk.main
