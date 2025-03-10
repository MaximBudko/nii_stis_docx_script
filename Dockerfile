# Используем образ Windows с поддержкой Ruby
FROM mcr.microsoft.com/windows/servercore:ltsc2022

# Устанавливаем Chocolatey для управления пакетами
RUN powershell -NoProfile -ExecutionPolicy Bypass -Command "& { \
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; \
    iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1')) \
}"

# Обновляем PATH для Chocolatey
ENV PATH="C:\\ProgramData\\chocolatey\\bin;${PATH}"

# Устанавливаем Ruby и GTK
RUN choco install -y ruby gtksharp

# Добавляем Ruby в PATH
ENV PATH="C:\\tools\\ruby31\\bin;${PATH}"

# Устанавливаем нужные гемы
RUN gem install bundler ocra gtk3

# Создаем директорию для приложения
WORKDIR /app

# Копируем файлы проекта
COPY . .

# Запускаем компиляцию в exe с ocra
RUN powershell -Command "ocra app.rb --windows --gem-all --output app.exe"

# Устанавливаем команду по умолчанию
CMD ["cmd"]
