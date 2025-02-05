require 'gtk3'
require 'roo'
require 'odf-report'
require 'docx'
require 'yaml'

CONFIG_FILE = "config.yml"

def load_config
  return {} unless File.exist?(CONFIG_FILE)
  YAML.load_file(CONFIG_FILE)
end

def save_config(config)
  File.write(CONFIG_FILE, config.to_yaml)
end

class ExcelToDocumentConverter
  def initialize
    @config = load_config
    @builder = Gtk::Builder.new
    @builder.add_from_file("ui/interface.glade")  # Загружаем интерфейс из файла
    @window = @builder.get_object("main_window")

    @excel_entry = @builder.get_object("excel_entry")
    @template_entry = @builder.get_object("template_entry")
    @format_combo = @builder.get_object("format_combo")

    @template_entry.text = @config["template"] if @config["template"]

    @builder.get_object("excel_button").signal_connect("clicked") { select_file(@excel_entry) }
    @builder.get_object("template_button").signal_connect("clicked") { select_file(@template_entry, save: true) }
    @builder.get_object("convert_button").signal_connect("clicked") { convert }

    @window.signal_connect("destroy") { Gtk.main_quit }
  end


  def select_file(entry, save: false)
    dialog = Gtk::FileChooserDialog.new(title: "Выберите файл", parent: @window, action: Gtk::FileChooserAction::OPEN)
    dialog.add_buttons(["Открыть", Gtk::ResponseType::ACCEPT], ["Отмена", Gtk::ResponseType::CANCEL])
    
    if dialog.run == Gtk::ResponseType::ACCEPT
      entry.text = dialog.filename
      @config["template"] = dialog.filename if save
      save_config(@config)
    end
    dialog.destroy
  end

  def convert
    excel_file = @excel_entry.text
    template_file = @template_entry.text
    format = @format_combo.active_text.downcase
    return unless File.exist?(excel_file) && File.exist?(template_file)
    
    if format == "odt"
      convert_excel_to_odt(excel_file, template_file)
    else
      convert_excel_to_docx(excel_file, template_file)
    end
  end

  def convert_excel_to_odt(excel_file, template_file)
    xlsx = Roo::Spreadsheet.open(excel_file)
    report = ODFReport::Report.new(template_file) do |r|
      r.add_table("TABLE", (1..xlsx.last_row).map { |row_num|
        (1..xlsx.last_column).map { |col_num| xlsx.cell(row_num, col_num).to_s }
      })
    end
    report.generate("converted.odt")
    puts "Конвертация завершена! Файл сохранен как converted.odt"
  end

  def convert_excel_to_docx(excel_file, template_file)
    doc = Docx::Document.open(template_file)
    xlsx = Roo::Spreadsheet.open(excel_file)
    
    xlsx.each_row_streaming do |row|
      doc.paragraphs << row.map(&:value).join("\t")
    end
    
    doc.save("converted.docx")
    puts "Конвертация завершена! Файл сохранен как converted.docx"
  end


  def run
    @window.show_all
    Gtk.main
  end
end

app = ExcelToDocumentConverter.new
app.run
