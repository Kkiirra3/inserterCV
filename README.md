# Генератор резюме на основе Google Docs

Этот проект представляет собой инструмент для автоматической генерации и обновления резюме на основе информации из Google Docs. Он объединяет данные из нескольких документов, вставляет информацию из JSON-файла и создает матрицу навыков на основе опыта работы.

## Структура проекта

```
/home/user/Desktop/insertedCVArch/
├───.gitignore
├───main.py
├───README.md
├───requirements.txt
├───.git/...
├───config/
│   ├───config.py
├───creds/
├───data/
├───src/
│   ├───__init__.py
│   ├───core/
│   │   ├───__init__.py
│   │   ├───document_processor.py
│   │   ├───skills_matrix_processor.py
│   │   ├───template_processor.py
│   ├───services/
│   │   ├───__init__.py
│   │   ├───google_service.py
│   └───utils/
│       ├───__init__.py
│       ├───formatting_utils.py
├───temp_docs/
```

- **`main.py`**: Главный скрипт, который запускает процесс генерации резюме.
- **`config/config.py`**: Файл конфигурации, содержащий все необходимые константы, такие как URL-адреса шаблонов, пути к файлам и настройки форматирования.
- **`data/`**: Каталог для хранения данных, таких как `template.json`.
- **`creds/`**: Каталог для хранения учетных данных для доступа к Google API.
- **`src/core/`**: Ядро приложения, содержащее основную логику.
    - **`document_processor.py`**: Отвечает за объединение документов, загрузку и выгрузку файлов, а также за вызов других обработчиков.
    - **`template_processor.py`**: Обрабатывает шаблоны, вставляя данные из `template.json` в документы.
    - **`skills_matrix_processor.py`**: Создает и обновляет матрицу навыков на основе данных о проектах.
- **`src/services/`**: Содержит сервисы для взаимодействия с внешними API.
    - **`google_service.py`**: Отвечает за аутентификацию и взаимодействие с Google Drive API.
- **`src/utils/`**: Содержит вспомогательные утилиты.
    - **`formatting_utils.py`**: Предоставляет функции для работы с форматированием в документах `.docx`.
- **`temp_docs/`**: Временный каталог для хранения загруженных и обработанных документов.

## Установка

1.  **Клонируйте репозиторий:**
    ```bash
    git clone git@github.com:Kkiirra3/inserterCV.git
    cd inserterCV
    ```

2.  **Создайте и активируйте виртуальное окружение:**
    ```bash
    python3 -m venv venv
    source venv/bin/activate
    ```

3.  **Установите зависимости:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Настройте учетные данные Google API:**
    - Перейдите в [Google Cloud Console](https://console.cloud.google.com/).
    - Создайте новый проект.
    - Включите **Google Drive API** и **Google Docs API**.
    - Создайте учетные данные типа **OAuth client ID** для **Desktop app**.
    - Скачайте JSON-файл с учетными данными и сохраните его как `creds/credentials.json`.

    При первом запуске скрипта вам будет предложено пройти аутентификацию в браузере. После этого будет создан файл `creds/token.pickle`, который будет использоваться для последующих запусков.

## Использование

1.  **Настройте `config/config.py`:**
    - Укажите правильные URL-адреса для `LISTPAGE_TEMPLATE_URL` и `MAIN_INFO_TEMPLATE_URL`.
    - При необходимости измените другие параметры, такие как `INPUT_SKILLS_DOC_ID` и настройки форматирования.

2.  **Заполните `data/template.json`:**
    - Внесите свои персональные данные, информацию о навыках, проектах и т.д.

3.  **Запустите скрипт:**
    ```bash
    python main.py
    ```

После выполнения скрипта в консоли появится ссылка на сгенерированный документ Google Docs.

## Как это работает

1.  **`main.py`** запускает `DocumentProcessor`.
2.  **`DocumentProcessor`** использует **`GoogleServiceManager`** для загрузки двух шаблонов Google Docs (`listpage` и `maininfo`) в виде файлов `.docx`.
3.  **`TemplateProcessor`** загружает данные из `data/template.json`.
4.  **`SkillsMatrixProcessor`** использует данные из `template.json` для создания матрицы навыков и сохраняет ее в отдельный `.docx` файл.
5.  **`TemplateProcessor`** вставляет данные из `template.json` в загруженные документы, заменяя плейсхолдеры (например, `{{NAME}}`, `{{TITLE}}`).
6.  **`DocumentProcessor`** объединяет обработанные документы в один файл.
7.  **`GoogleServiceManager`** загружает итоговый документ на Google Drive и преобразует его в формат Google Docs.
8.  Скрипт выводит ссылку на созданный документ.
