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


class FileChooserApp < Gtk::Window
  SETTINGS_FILE = File.expand_path('saves/settings.json', __dir__)
  FIELD_LABELS = ["Перв. примен.", "Разраб.", "Пров.", "Н. контр.", "Утв.", "Дец. номер", "Наименование устройства", "Наименование организации","Номер изменения", "Нов / Зам", "Номер извещения"]
  FIELD_LABELS_FOR_REMOVED = ["perv_primen", "razrab", "prover", "n_kontr", "utverd", "blpa", "device_name", "company_name","n_i", "n_z", "nom_iz"]
 
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

    FIELD_LABELS.each_with_index do |label_text, i|
      hbox_label = Gtk::Box.new(:horizontal, 10)
      main_box.pack_start(hbox_label, expand: false, fill: false, padding: 0)
      
      label = Gtk::Label.new(label_text)
      label.set_xalign(0)
      hbox_label.pack_start(label, expand: true, fill: true, padding: 0)
      
      hbox = Gtk::Box.new(:horizontal, 10)
      main_box.pack_start(hbox, expand: false, fill: false, padding: 0)

      check_button = Gtk::CheckButton.new
      check_button.active = @settings["checkbox_#{i}"]
      hbox.pack_start(check_button, expand: false, fill: false, padding: 0)
      
      entry = Gtk::Entry.new
      entry.set_hexpand(true)
      entry.set_text(@settings["checkbox_#{i}"] ? (@settings["field_#{i}"] || "") : "")
      hbox.pack_start(entry, expand: true, fill: true, padding: 0)
      
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
    end

    @convert_button = Gtk::Button.new(label: "Сконвертировать")
    main_box.pack_end(@convert_button, expand: false, fill: false, padding: 0)
    @convert_button.signal_connect("clicked") { on_convert_clicked }
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
  def load_settings
    return {} unless File.exist?(SETTINGS_FILE)
    JSON.parse(File.read(SETTINGS_FILE))
  rescue
    {}
  end

  def save_settings
    File.write(SETTINGS_FILE, JSON.pretty_generate(@settings))
  end
end

if __FILE__ == $0
  app = FileChooserApp.new
  app.show_all
  Gtk.main
end
