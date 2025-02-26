require 'gtk3'
require 'json'

class FileChooserApp < Gtk::Window
  SETTINGS_FILE = "saves/settings.json"
  FIELD_LABELS = ["Перв. примен.", "Разраб.", "Пров.", "Н. контр.", "Утв.", "Дец. номер", "Наименование устройства", "Наименование организации", "Литера"]
  FIELD_LABELS_FOR_REMOVED = ["perv_primen", "razrab", "prover", "n_kontr", "utverd", "blpa", "device_name", "company_name", "litera"]
 
 
  def initialize
    super(Gtk::WindowType::TOPLEVEL)
    set_title("Приложение")
    set_default_size(500, 500)
    set_window_position(Gtk::WindowPosition::CENTER)
    signal_connect("destroy") { Gtk.main_quit }
    
    @settings = load_settings
    @checkbox_states = {}
    @save_file_path = ""
    @save_directory = ""

    vbox = Gtk::Box.new(:vertical, 10)
    vbox.set_margin_top(10)
    vbox.set_margin_bottom(10)
    vbox.set_margin_start(10)
    vbox.set_margin_end(10)
    add(vbox)

    hbox_file = Gtk::Box.new(:horizontal, 10)
    vbox.pack_start(hbox_file, expand: false, fill: false, padding: 0)

    @entry = Gtk::Entry.new
    @entry.set_hexpand(true)
    @entry.set_editable(false)
    hbox_file.pack_start(@entry, expand: true, fill: true, padding: 0)

    @button = Gtk::Button.new(label: "Выбрать файл")
    @button.signal_connect("clicked") { on_file_clicked }
    hbox_file.pack_start(@button, expand: false, fill: false, padding: 0)

    @text_entries = {}
    @check_buttons = {}

    FIELD_LABELS.each_with_index do |label_text, i|
      hbox_label = Gtk::Box.new(:horizontal, 10)
      vbox.pack_start(hbox_label, expand: false, fill: false, padding: 0)
      
      label = Gtk::Label.new(label_text)
      label.set_xalign(0)
      hbox_label.pack_start(label, expand: true, fill: true, padding: 0)
      
      hbox = Gtk::Box.new(:horizontal, 10)
      vbox.pack_start(hbox, expand: false, fill: false, padding: 0)

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
    vbox.pack_end(@convert_button, expand: false, fill: false, padding: 0)
    @convert_button.signal_connect("clicked") { on_convert_clicked }
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
    end
    dialog.destroy
  end

  def on_convert_clicked
    excel_path = @entry.text
    if excel_path.empty?
      puts "Файл не выбран!"
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
      require_relative 'index'
      ExcelToDocx.generate_docx(docx_path="template/shablon_pr.docx", excel_path, field_values, @save_file_path)
    end
    save_dialog.destroy
  end

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
