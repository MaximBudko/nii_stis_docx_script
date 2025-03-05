FROM ubuntu:latest

# Устанавливаем зависимости
RUN apt update && apt install -y \
    mingw-w64 \
    ruby-full \
    ruby-dev \
    make \
    gcc \
    zip \
    unzip \
    git \
    wget \
    pkg-config \
    libgtk-3-dev

# Устанавливаем Bundler
RUN gem install bundler

# Устанавливаем необходимые гемы
COPY Gemfile Gemfile.lock ./
RUN bundle install

# Устанавливаем Ocra
RUN gem install ocra

# Копируем исходный код
WORKDIR /app
COPY . /app

# Скачиваем GTK3 для Windows
RUN wget -O gtk.zip "https://download.gnome.org/binaries/win64/gtk+-3.24.24.zip" \
    && unzip gtk.zip -d /usr/local/gtk \
    && rm gtk.zip

# Экспортируем переменные среды для сборки
ENV PATH="/usr/local/gtk/bin:$PATH"
ENV PKG_CONFIG_PATH="/usr/local/gtk/lib/pkgconfig"

# Компиляция EXE с Ocra
CMD ["sh", "-c", "x86_64-w64-mingw32-ruby -S ocra app.rb --add-all-core --gem-all --dll gtk3"]
