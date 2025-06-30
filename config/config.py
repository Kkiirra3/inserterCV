import os

class Config:
    """
    Класс для хранения конфигурационных данных и констант.
    """
    # --- ID и пути к файлам ---
    INPUT_SKILLS_DOC_ID = '1cQI3Ve289uae_EYFmiz6AEgZDHz1UqBehs4hGTsyXyE'
    OUTPUT_SKILLS_DOCX = 'temp_docs/updated_skills_matrix.docx'
    TEMPLATE_JSON = 'data/template.json'
    
    # --- Учетные данные API ---
    CREDENTIALS_JSON = 'creds/credentials.json'
    TOKEN_PICKLE = 'creds/token.pickle'
    SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']
    
    # --- Настройки форматирования таблицы в матрице ---
    BORDER_COLOR = "C63031"
    BORDER_SIZE = "4"

    # Template URLs from environment variables
    LISTPAGE_TEMPLATE_URL = 'https://docs.google.com/document/d/1W_OfVU_G8drgW5cR5DsAK9x5ZznGbznNy--MZOk9Cjw/edit?tab=t.0'
    MAIN_INFO_TEMPLATE_URL = 'https://docs.google.com/document/d/1grLACx3VSGdEuD0VkhpYuRoAHRmzHXXqls73hKhZP_o/edit?tab=t.0#heading=h.tvxmemvzgy6z'
