# ENV['GI_TYPELIB_PATH'] = File.expand_path('for_build\girepository-1.0', __dir__)
# ENV['FONTCONFIG_PATH'] = File.expand_path('for_build\fonts', __dir__)
# ENV['XDG_DATA_DIRS'] = File.expand_path('for_build\share', __dir__)
# ENV['GSETTINGS_SCHEMA_DIR'] = File.expand_path('for_build\schemas', __dir__)

require 'gtk3'
require 'json'
require 'roo'
require 'write_xlsx'
require 'rubyXL'
require 'zip'
require 'nokogiri'
require 'fileutils'
require 'pp'
require 'stringio'
require_relative 'utils/category_manager'


class FileChooserApp < Gtk::Window
  SETTINGS_FILE = File.expand_path('saves/settings.json', __dir__)
  FIELD_LABELS = ["Перв. примен.", "Разраб.", "Пров.", "Н. контр.", "Утв.", "Дец. номер", "Наименование устройства", "Наименование организации","Номер изменения", "Нов / Зам", "Номер извещения"]
  FIELD_LABELS_FOR_REMOVED = ["perv_primen", "razrab", "prover", "n_kontr", "utverd", "blpa", "device_name", "company_name","n_i", "n_z", "nom_iz"]
  DEFAULT_SPEC_ITER = 1
  DEFAULT_SPEC_ITER2 = 30

  def initialize
    super(Gtk::WindowType::TOPLEVEL)
    set_title("Приложение")
    set_default_size(600, 500)
    set_window_position(Gtk::WindowPosition::CENTER)
    signal_connect("destroy") { Gtk.main_quit }
    
    @settings = load_settings
    @checkbox_states = {}
    @save_file_path = ""
    @save_directory = ""
    @spec_iter = @settings['spec_iter'] || DEFAULT_SPEC_ITER
    @spec_iter2 = @settings['spec_iter2'] || DEFAULT_SPEC_ITER2

    notebook = Gtk::Notebook.new
    add(notebook)

    main_box = Gtk::Box.new(:vertical, 10)
    main_box.set_margin_top(10)
    main_box.set_margin_bottom(10)
    main_box.set_margin_start(10)
    main_box.set_margin_end(10)
    notebook.append_page(main_box, Gtk::Label.new("Главная"))

  #----------------Код вкладки Compressed-------------------

    compressed_box = Gtk::Box.new(:vertical, 10)
    compressed_box.set_margin_top(10)
    compressed_box.set_margin_bottom(10)
    compressed_box.set_margin_start(10)
    compressed_box.set_margin_end(10)
    notebook.append_page(compressed_box, Gtk::Label.new("Compressed"))
    
  #---------------------------------------------------------

  #-------------Код вкладки дополнительные настройки--------
  additional_settings = Gtk::Box.new(:vertical, 10)
  additional_settings.set_margin_top(10)
  additional_settings.set_margin_bottom(10)
  additional_settings.set_margin_start(10)
  additional_settings.set_margin_end(10)
  notebook.append_page(additional_settings, Gtk::Label.new("Доп. настройки"))

  # Добавляем метку и текстовое поле для spec_iter
  spec_iter_label = Gtk::Label.new("Подсчет элементов в <<Спецификации>> начинается с:")
  additional_settings.pack_start(spec_iter_label, expand: false, fill: false, padding: 5)

  spec_iter_entry = Gtk::Entry.new
  spec_iter_entry.set_hexpand(true)
  spec_iter_entry.set_text(@spec_iter.to_s)
  additional_settings.pack_start(spec_iter_entry, expand: false, fill: false, padding: 5)

  # Ограничиваем ввод только числами
  spec_iter_entry.signal_connect("changed") do
    text = spec_iter_entry.text
    if text.empty?
      @spec_iter = DEFAULT_SPEC_ITER
    elsif text.match?(/^\d+$/)
      @spec_iter = text.to_i
    else
      spec_iter_entry.set_text(@spec_iter.to_s)
    end
    @settings['spec_iter'] = @spec_iter
    save_settings
  end
  #---------------------------------------------------------
  #-----------------------Вторая настройка------------------
  spec_iter_label2 = Gtk::Label.new("Ограничения длинны наименования в <<Спецификации>> с:")
  additional_settings.pack_start(spec_iter_label2, expand: false, fill: false, padding: 5)

  spec_iter_entry2 = Gtk::Entry.new
  spec_iter_entry2.set_hexpand(true)
  spec_iter_entry2.set_text(@spec_iter2.to_s)
  additional_settings.pack_start(spec_iter_entry2, expand: false, fill: false, padding: 5)

  # Ограничиваем ввод только числами
  spec_iter_entry2.signal_connect("changed") do
    text = spec_iter_entry2.text
    if text.empty?
      @spec_iter2 = DEFAULT_SPEC_ITER2
    elsif text.match?(/^\d+$/)
      @spec_iter2 = text.to_i
    else
      spec_iter_entry2.set_text(@spec_iter2.to_s)
    end
    @settings['spec_iter2'] = @spec_iter2
    save_settings
  end
#---------------------------------------------------------

#-----------------------Код редактирования словарей-------
  # Добавляем разделитель
  category_separator = Gtk::Separator.new(:horizontal)
  additional_settings.pack_start(category_separator, expand: false, fill: true, padding: 10)

  # Заголовок секции категорий
  category_label = Gtk::Label.new("Настройки категорий компонентов")
  category_label.set_xalign(0)
  additional_settings.pack_start(category_label, expand: false, fill: false, padding: 5)

  # Создаем поля для ввода новой категории
  input_grid = Gtk::Grid.new
  input_grid.set_column_spacing(5)
  input_grid.set_row_spacing(5)
  additional_settings.pack_start(input_grid, expand: false, fill: true, padding: 5)

  # Поле для ключа категории
  key_label = Gtk::Label.new("Ключ:")
  key_entry = Gtk::Entry.new
  key_entry.set_width_chars(10)
  input_grid.attach(key_label, 0, 0, 1, 1)
  input_grid.attach(key_entry, 1, 0, 1, 1)

  # Поле для единственного числа
  singular_label = Gtk::Label.new("Ед. число:")
  singular_entry = Gtk::Entry.new
  singular_entry.set_width_chars(20)
  input_grid.attach(singular_label, 2, 0, 1, 1)
  input_grid.attach(singular_entry, 3, 0, 1, 1)

  # Поле для множественного числа
  plural_label = Gtk::Label.new("Мн. число:")
  plural_entry = Gtk::Entry.new
  plural_entry.set_width_chars(20)
  input_grid.attach(plural_label, 4, 0, 1, 1)
  input_grid.attach(plural_entry, 5, 0, 1, 1)



  # Создаем список категорий
  scrolled_window = Gtk::ScrolledWindow.new
  scrolled_window.set_policy(:automatic, :automatic)
  scrolled_window.set_size_request(-1, 200)
  additional_settings.pack_start(scrolled_window, expand: true, fill: true, padding: 5)

  category_store = Gtk::ListStore.new(String, String, String)
  category_view = Gtk::TreeView.new(category_store)

  # Добавляем колонки
  renderer = Gtk::CellRendererText.new
  column1 = Gtk::TreeViewColumn.new("Ключ", renderer, text: 0)
  column2 = Gtk::TreeViewColumn.new("Ед. число", renderer, text: 1)
  column3 = Gtk::TreeViewColumn.new("Мн. число", renderer, text: 2)

  category_view.append_column(column1)
  category_view.append_column(column2)
  category_view.append_column(column3)

  scrolled_window.add(category_view)
  # Кнопки управления
  button_box = Gtk::Box.new(:horizontal, 5)
  additional_settings.pack_start(button_box, expand: false, fill: true, padding: 5)
  
  add_button = Gtk::Button.new(label: "Добавить")
  update_button = Gtk::Button.new(label: "Обновить")
  delete_button = Gtk::Button.new(label: "Удалить")
  
  button_box.pack_start(add_button, expand: true, fill: true, padding: 2)
  button_box.pack_start(update_button, expand: true, fill: true, padding: 2)
  button_box.pack_start(delete_button, expand: true, fill: true, padding: 2)
  
  # Инициализируем список категорий
  update_category_list(category_store)
  
  # Настраиваем обработчики событий кнопок
  setup_category_buttons(add_button, update_button, delete_button, key_entry, singular_entry, plural_entry, category_store)
  
  # Настраиваем выбор категории из списка
  category_view.signal_connect("cursor-changed") do
    selection = category_view.selection
    if iter = selection.selected
      key_entry.text = iter[0]
      singular_entry.text = iter[1]
      plural_entry.text = iter[2]
    end
  end

  update_category_list(category_store)

  # Обработчики событий
  add_button.signal_connect("clicked") do
    key = key_entry.text.strip
    singular = singular_entry.text.strip
    plural = plural_entry.text.strip

    if key.empty? || singular.empty? || plural.empty?
      log_message("Ошибка: Все поля должны быть заполнены")
    else
      if CategoryManager.add_category(key, singular, plural)
        update_category_list(category_store)
        log_message("Категория успешно добавлена")
        key_entry.text = ""
        singular_entry.text = ""
        plural_entry.text = ""
      else
        log_message("Ошибка: Категория с таким ключом уже существует")
      end
    end
  end

  update_button.signal_connect("clicked") do
    key = key_entry.text.strip
    singular = singular_entry.text.strip
    plural = plural_entry.text.strip

    if key.empty? || singular.empty? || plural.empty?
      log_message("Ошибка: Все поля должны быть заполнены")
    else
      if CategoryManager.update_category(key, singular, plural)
        update_category_list(category_store)
        log_message("Категория успешно обновлена")
      else
        log_message("Ошибка: Категория не найдена")
      end
    end
  end

  delete_button.signal_connect("clicked") do
    key = key_entry.text.strip
    if key.empty?
      log_message("Ошибка: Введите ключ категории для удаления")
    else
      if CategoryManager.remove_category(key)
        update_category_list(category_store)
        log_message("Категория успешно удалена")
        key_entry.text = ""
        singular_entry.text = ""
        plural_entry.text = ""
      else
        log_message("Ошибка: Категория не найдена")
      end
    end
  end

  # Выбор категории из списка
  category_view.signal_connect("cursor-changed") do
    selection = category_view.selection
    if iter = selection.selected
      key_entry.text = iter[0]
      singular_entry.text = iter[1]
      plural_entry.text = iter[2]
    end
  end
  #---------------------------------------------------------
    @log_textview = Gtk::TextView.new
    @log_textview.editable = false
    @log_buffer = @log_textview.buffer
    log_scrolled = Gtk::ScrolledWindow.new
    log_scrolled.set_policy(:automatic, :automatic)
    log_scrolled.add(@log_textview)
    notebook.append_page(log_scrolled, Gtk::Label.new("Логирование"))

  #--------------------Код выбора файла в main-----------------------
    hbox_file = Gtk::Box.new(:horizontal, 10)
    main_box.pack_start(hbox_file, expand: false, fill: false, padding: 0)

    @entry = Gtk::Entry.new
    @entry.set_hexpand(true)
    @entry.set_editable(true)
    hbox_file.pack_start(@entry, expand: true, fill: true, padding: 0)

    @button = Gtk::Button.new(label: "Выбрать файл")
    @button.signal_connect("clicked") { on_file_clicked }
    hbox_file.pack_start(@button, expand: false, fill: false, padding: 0)

  #-------------------------------------------------------------------------

  #---------------Код выбора файла в Compressed-----------------------------
    set_path_file_comp = Gtk::Box.new(:horizontal, 10)
    compressed_box.pack_start(set_path_file_comp, expand: false, fill: false, padding: 0)

    @entry_compressed = Gtk::Entry.new
    @entry_compressed.set_hexpand(true)
    @entry_compressed.set_editable(true)
    set_path_file_comp.pack_start(@entry_compressed, expand: true, fill: true, padding: 0)


    @button_set_compressed = Gtk::Button.new(label: "Выбрать файл")
    @button_set_compressed.signal_connect("clicked") {compresed_bom_controller}
    set_path_file_comp.pack_start(@button_set_compressed, expand: false, fill: false, padding: 0)
  #-------------------------------------------------------------------------

  #----------------Кнопка конвертации Compressed----------------------------
    @compressed_bom_button = Gtk::Button.new(label: "Compressed")
    compressed_box.pack_start(@compressed_bom_button, expand: false, fill: false, padding: 0)
    @compressed_bom_button.signal_connect("clicked") { compressed_bom_button_clicked }
  #-------------------------------------------------------------------------

  @text_entries = {}
  @check_buttons = {}

  # Создаем основной контейнер для двух колонок
  columns_box = Gtk::Box.new(:horizontal, 20)
  main_box.pack_start(columns_box, expand: true, fill: true, padding: 10)

  # Создаем левую и правую колонки
  left_column = Gtk::Box.new(:vertical, 5)
  right_column = Gtk::Box.new(:vertical, 5)
  
  # Добавляем вертикальный разделитель между колонками
  separator = Gtk::Separator.new(:vertical)
  
  # Упаковываем колонки и разделитель
  columns_box.pack_start(left_column, expand: true, fill: true, padding: 10)
  columns_box.pack_start(separator, expand: false, fill: true, padding: 0)
  columns_box.pack_start(right_column, expand: true, fill: true, padding: 10)

  # Создаем контейнер для последнего поля под колонками
  bottom_field_box = Gtk::Box.new(:vertical, 5)
  main_box.pack_start(bottom_field_box, expand: false, fill: true, padding: 10)

  @text_entries = {}
  @check_buttons = {}

  # Вычисляем количество полей для каждой колонки
  fields_count = FIELD_LABELS.length
  fields_per_column = (fields_count - 1) / 2
  last_field_index = fields_count - 1

  FIELD_LABELS.each_with_index do |label_text, i|
    # Определяем, куда помещать текущее поле
    current_container = if i == last_field_index
      bottom_field_box  # Последнее поле идет в нижний контейнер
    else
      i < fields_per_column ? left_column : right_column
    end
    
    # Создаем контейнер для метки и поля
    field_box = Gtk::Box.new(:vertical, 5)
    current_container.pack_start(field_box, expand: false, fill: true, padding: 5)
    
    # Метка
    label = Gtk::Label.new(label_text)
    label.set_xalign(0)
    field_box.pack_start(label, expand: false, fill: false, padding: 0)
    
    # Контейнер для чекбокса и поля ввода
    input_box = Gtk::Box.new(:horizontal, 5)
    field_box.pack_start(input_box, expand: true, fill: true, padding: 0)

    # Чекбокс
    check_button = Gtk::CheckButton.new
    check_button.active = @settings["checkbox_#{i}"]
    input_box.pack_start(check_button, expand: false, fill: false, padding: 0)
    
    # Поле ввода
    entry = Gtk::Entry.new
    entry.set_hexpand(true)
    entry.set_text(@settings["checkbox_#{i}"] ? (@settings["field_#{i}"] || "") : "")
    input_box.pack_start(entry, expand: true, fill: true, padding: 0)
    
    @checkbox_states[i] = check_button.active?
    
    check_button.signal_connect("toggled") do
      @checkbox_states[i] = check_button.active?
      @settings["checkbox_#{i}"] = check_button.active?
      if check_button.active?
        @settings["field_#{i}"] = entry.text
      else
        @settings.delete("field_#{i}")
        entry.set_text("")
      end
      save_settings
    end
    
    entry.signal_connect("changed") do
      if check_button.active?
        @settings["field_#{i}"] = entry.text
        save_settings
      end
    end
    
    @text_entries[label_text] = entry
    @check_buttons[label_text] = check_button

    # Добавляем горизонтальный разделитель после каждого поля
    if i != last_field_index # Не добавляем разделитель после последнего поля
      current_container.pack_start(Gtk::Separator.new(:horizontal), expand: false, fill: true, padding: 2)
    end
  end

  button_box = Gtk::Box.new(:horizontal, 10)
  button_box.homogeneous = true
  
  @convert_button_vedomost = Gtk::Button.new(label: "Ведомость")
  @convert_button_spec = Gtk::Button.new(label: "Спецификация")
  @convert_button = Gtk::Button.new(label: "Перечень")
  
  button_box.pack_start(@convert_button_vedomost, expand: true, fill: true, padding: 2)
  button_box.pack_start(@convert_button_spec, expand: true, fill: true, padding: 2)
  button_box.pack_start(@convert_button, expand: true, fill: true, padding: 2)
  
  @convert_button_vedomost.signal_connect("clicked") { on_vedomost_clicked }
  @convert_button_spec.signal_connect("clicked") { specifiacation_button_clicked }
  @convert_button.signal_connect("clicked") { on_convert_clicked }
  
  main_box.pack_end(button_box, expand: true, fill: true, padding: 2)

end

def log_message(message)
  iter = @log_buffer.end_iter
  @log_buffer.insert(iter, "#{Time.now}: #{message}\n")
end

def on_file_clicked
  dialog = Gtk::FileChooserDialog.new(
    title: "Выберите файл",
    parent: self,
    action: Gtk::FileChooserAction::OPEN,
    buttons: [["Отмена", Gtk::ResponseType::CANCEL], ["Открыть", Gtk::ResponseType::OK]]
  )

  if dialog.run == Gtk::ResponseType::OK
    @entry.set_text(dialog.filename)
    log_message("Файл выбран: #{dialog.filename}")
  end
  dialog.destroy
end

def compresed_bom_controller
  dialog_comp = Gtk::FileChooserDialog.new(
    title: "Выберите файл",
    parent: self,
    action: Gtk::FileChooserAction::OPEN,
    buttons: [["Отмена", Gtk::ResponseType::CANCEL], ["Открыть", Gtk::ResponseType::OK]]
  )

  if dialog_comp.run == Gtk::ResponseType::OK
    @entry_compressed.set_text(dialog_comp.filename)
    log_message("Файл выбран: #{dialog_comp.filename}")
  end
  dialog_comp.destroy
end


def on_convert_clicked
  excel_path = @entry.text
  if excel_path.empty?
    log_message("Ошибка: Файл не выбран!")
    return
  end

  save_dialog = Gtk::FileChooserDialog.new(
    title: "Сохранить файл",
    parent: self,
    action: Gtk::FileChooserAction::SAVE,
    buttons: [["Отмена", Gtk::ResponseType::CANCEL], ["Сохранить", Gtk::ResponseType::OK]]
  )
  save_dialog.set_do_overwrite_confirmation(true)

  if save_dialog.run == Gtk::ResponseType::OK
    @save_file_path = save_dialog.filename
    file_name = File.basename(@save_file_path)
    field_values = Hash[FIELD_LABELS_FOR_REMOVED.zip(@text_entries.values.map(&:text))]
    begin
      require_relative 'index'
      path_to_converted_docx = File.expand_path('template/shablon_pr.docx', __dir__)
      ExcelToDocx.generate_docx(path_to_converted_docx, excel_path, field_values, @save_file_path)
      log_message("Файл успешно сконвертирован: #{@save_file_path}")
    rescue StandardError => e
      log_message("Ошибка конвертации: #{e.message}")
    end
  end
  save_dialog.destroy
end

#----------------------------------Вызов спецификации---------------------------
def specifiacation_button_clicked
  excel_path = @entry.text
  if excel_path.empty?
    log_message("Ошибка: Файл не выбран!")
    return
  end

  save_dialog = Gtk::FileChooserDialog.new(
    title: "Сохранить файл",
    parent: self,
    action: Gtk::FileChooserAction::SAVE,
    buttons: [["Отмена", Gtk::ResponseType::CANCEL], ["Сохранить", Gtk::ResponseType::OK]]
  )
  save_dialog.set_do_overwrite_confirmation(true)

  if save_dialog.run == Gtk::ResponseType::OK
    @save_file_path = save_dialog.filename
    file_name = File.basename(@save_file_path)
    field_values = Hash[FIELD_LABELS_FOR_REMOVED.zip(@text_entries.values.map(&:text))]
    input_int = @spec_iter
    inupt_del = @spec_iter2
    begin
      require_relative 'spec'
      path_to_converted_docx = File.expand_path('template/shablon_sp.docx', __dir__)
      Spec.generate_spec(path_to_converted_docx, excel_path, field_values, @save_file_path, input_int, inupt_del)
      log_message("Файл успешно сконвертирован: #{@save_file_path}")
    rescue StandardError => e
      log_message("Ошибка конвертации: #{e.message}")
    end
  end
  save_dialog.destroy
end
#--------------------------------------------------------------
#-----------------------Compressed BOM вызов функции----------

  def compressed_bom_button_clicked
    excel_path = @entry_compressed.text
    if excel_path.empty?
      log_message("Ошибка: Файл не выбран!")
      return
    end

    save_dialog = Gtk::FileChooserDialog.new(
      title: "Сохранить файл",
      parent: self,
      action: Gtk::FileChooserAction::SAVE,
      buttons: [["Отмена", Gtk::ResponseType::CANCEL], ["Сохранить", Gtk::ResponseType::OK]]
    )
    save_dialog.set_do_overwrite_confirmation(true)

    if save_dialog.run == Gtk::ResponseType::OK
      @save_file_path = save_dialog.filename
      file_name = File.basename(@save_file_path)
      begin
        require_relative 'compressed_bom'
        CompressedBom.process_excel(excel_path, "#{@save_file_path}.xlsx")
        log_message("Файл успешно сконвертирован: #{@save_file_path}")
      rescue StandardError => e
        log_message("Ошибка конвертации: #{e.message}")
      end
    end
    save_dialog.destroy
  end
  #-------------------------------------------------------------

  #----------------------Vedomst вызов функций -----------------
  def on_vedomost_clicked
    excel_path = @entry.text
    if excel_path.empty?
      log_message("Ошибка: Файл не выбран!")
      return
    end

    save_dialog = Gtk::FileChooserDialog.new(
      title: "Сохранить файл",
      parent: self,
      action: Gtk::FileChooserAction::SAVE,
      buttons: [["Отмена", Gtk::ResponseType::CANCEL], ["Сохранить", Gtk::ResponseType::OK]]
    )
    save_dialog.set_do_overwrite_confirmation(true)

    if save_dialog.run == Gtk::ResponseType::OK
      @save_file_path = save_dialog.filename
      file_name = File.basename(@save_file_path)
      field_values = Hash[FIELD_LABELS_FOR_REMOVED.zip(@text_entries.values.map(&:text))]
      begin
        require_relative 'vedomost'
        path_to_converted_docx = File.expand_path('template/shablon_vp.docx', __dir__)
        Vedomost.generate_docx(path_to_converted_docx, excel_path, field_values, @save_file_path)
        log_message("Файл успешно сконвертирован: #{@save_file_path}")
      rescue StandardError => e
        log_message("Ошибка конвертации: #{e.message}")
      end
    end
    save_dialog.destroy
  end
#-------------------------------------------------------------


  def load_settings
    return {} unless File.exist?(SETTINGS_FILE)
    JSON.parse(File.read(SETTINGS_FILE))
  rescue
    {}
  end

  def save_settings
    File.write(SETTINGS_FILE, JSON.pretty_generate(@settings))
  end

  private

  def update_category_list(store)
    store.clear
    categories = CategoryManager.load_categories
    categories.each do |key, values|
      iter = store.append
      iter[0] = key
      iter[1] = values[0]
      iter[2] = values[1]
    end
  end

  def add_category(key, singular, plural)
    return false if key.empty? || singular.empty? || plural.empty?
    CategoryManager.add_category(key, singular, plural)
  end

  def remove_category(key)
    return false if key.empty?
    CategoryManager.remove_category(key)
  end

  def update_category(key, singular, plural)
    return false if key.empty? || singular.empty? || plural.empty?
    CategoryManager.update_category(key, singular, plural)
  end

  # Обновим обработчики событий кнопок
  def setup_category_buttons(add_button, update_button, delete_button, key_entry, singular_entry, plural_entry, category_store)
    add_button.signal_connect("clicked") do
      key = key_entry.text.strip
      singular = singular_entry.text.strip
      plural = plural_entry.text.strip

      if add_category(key, singular, plural)
        update_category_list(category_store)
        log_message("Категория успешно добавлена")
        key_entry.text = ""
        singular_entry.text = ""
        plural_entry.text = ""
      else
        log_message("Ошибка: Все поля должны быть заполнены или категория уже существует")
      end
    end

    update_button.signal_connect("clicked") do
      key = key_entry.text.strip
      singular = singular_entry.text.strip
      plural = plural_entry.text.strip

      if update_category(key, singular, plural)
        update_category_list(category_store)
        log_message("Категория успешно обновлена")
      else
        log_message("Ошибка: Все поля должны быть заполнены или категория не найдена")
      end
    end

    delete_button.signal_connect("clicked") do
      key = key_entry.text.strip
      if remove_category(key)
        update_category_list(category_store)
        log_message("Категория успешно удалена")
        key_entry.text = ""
        singular_entry.text = ""
        plural_entry.text = ""
      else
        log_message("Ошибка: Введите ключ категории для удаления")
      end
    end
  end
end

if __FILE__ == $0
app = FileChooserApp.new
app.show_all
Gtk.main
end
