require 'gtk3'

class FileChooserApp < Gtk::Window
  def initialize
    super(Gtk::WindowType::TOPLEVEL)
    set_title("GTK3 FileChooser with GtkFixed")
    set_default_size(400, 150)

    fixed = Gtk::Fixed.new
    add(fixed)

    @entry = Gtk::Entry.new
    @entry.set_size_request(250, 40)
    @entry.set_editable(false)
    fixed.put(@entry, 20, 20)

    @button = Gtk::Button.new(label: "Выбрать файл")
    @button.set_size_request(120, 40)
    @button.signal_connect("clicked") { on_file_clicked }
    fixed.put(@button, 280, 20)

    @convert_button = Gtk::Button.new(label: "Сконвертировать")
    @convert_button.set_size_request(380, 40)
    @convert_button.signal_connect("clicked") { on_convert_clicked }
    fixed.put(@convert_button, 20, 80)
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
    file_path = @entry.text
    if file_path.empty?
      puts "Файл не выбран!"
    else
      require_relative 'converter_script'
      ConverterScript.run(file_path)
    end
  end
end

if __FILE__ == $0
  app = FileChooserApp.new
  app.signal_connect("destroy") { Gtk.main_quit }
  app.show_all
  Gtk.main
end
