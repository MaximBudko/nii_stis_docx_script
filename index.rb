require 'gtk3'
require 'roo'
require 'docx'

class ExcelToWordConverter
  def initialize
    @builder = Gtk::Builder.new
    @builder.add_from_string(interface)
    @window = @builder.get_object("main_window")
    
    @excel_entry = @builder.get_object("excel_entry")
    @word_entry = @builder.get_object("word_entry")
    
    @builder.get_object("excel_button").signal_connect("clicked") { select_file(@excel_entry) }
    @builder.get_object("word_button").signal_connect("clicked") { select_file(@word_entry) }
    @builder.get_object("convert_button").signal_connect("clicked") { convert }
    
    @window.signal_connect("destroy") { Gtk.main_quit }
  end

  def select_file(entry)
    dialog = Gtk::FileChooserDialog.new(title: "Выберите файл", parent: @window, action: Gtk::FileChooserAction::OPEN)
    dialog.add_buttons(["Открыть", Gtk::ResponseType::ACCEPT], ["Отмена", Gtk::ResponseType::CANCEL])
    
    if dialog.run == Gtk::ResponseType::ACCEPT
      entry.text = dialog.filename
    end
    dialog.destroy
  end

  def convert
    excel_file = @excel_entry.text
    word_file = @word_entry.text
    return unless File.exist?(excel_file) && File.exist?(word_file)
    
    convert_excel_to_word(excel_file, word_file)
  end

  def convert_excel_to_word(excel_file, word_file)
    xlsx = Roo::Spreadsheet.open(excel_file)
    doc = Docx::Document.open(word_file)
    
    table_data = []
    (1..xlsx.last_row).each do |row_num|
      row_data = (1..xlsx.last_column).map { |col_num| xlsx.cell(row_num, col_num).to_s }
      table_data << row_data
    end
    
    doc.paragraphs << "\n"
    table = doc.add_table(table_data.size, table_data[0].size)
    
    table_data.each_with_index do |row, i|
      row.each_with_index do |cell, j|
        table[i][j].text = cell
      end
    end
    
    doc.save(word_file)
    puts "Конвертация завершена!"
  end

  def interface
    <<-UI
      <interface>
        <object class='GtkWindow' id='main_window'>
          <property name='title'>Excel to Word Converter</property>
          <child>
            <object class='GtkBox'>
              <property name='orientation'>vertical</property>
              <child>
                <object class='GtkEntry' id='excel_entry'/>
              </child>
              <child>
                <object class='GtkButton' id='excel_button'>
                  <property name='label'>Обзор</property>
                </object>
              </child>
              <child>
                <object class='GtkEntry' id='word_entry'/>
              </child>
              <child>
                <object class='GtkButton' id='word_button'>
                  <property name='label'>Обзор</property>
                </object>
              </child>
              <child>
                <object class='GtkButton' id='convert_button'>
                  <property name='label'>Конвертировать</property>
                </object>
              </child>
            </object>
          </child>
        </object>
      </interface>
    UI
  end

  def run
    @window.show_all
    Gtk.main
  end
end

app = ExcelToWordConverter.new
app.run
