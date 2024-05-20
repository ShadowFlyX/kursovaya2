def read_settings(file_path: str) -> dict:

    settings = {}
    with open(file_path, 'r') as file:
        for line in file:
            line = line.strip()
            if line and not line.startswith('#'):
                key, value = line.split('=', 1)
                settings[key.strip()] = value.strip()
    return settings

def create_connection_string(settings):
    
    server = settings.get('SERVER')
    database = settings.get('DATABASE')
    driver = settings.get('DRIVER', 'ODBC+Driver+18+for+SQL+Server')
    
    # Определяем параметры для подключения
    username = settings.get('USERNAME')
    password = settings.get('PASSWORD')
    trusted_connection = settings.get('TRUSTED_CONNECTION', 'YES')
    trustservercertificate = settings.get('TRUSTSERVERCERTIFICATE', 'YES')
    
    if username and password:
        # Если указаны логин и пароль, используем их
        connection_string = f"mssql+pyodbc://{username}:{password}@{server}/{database}?"
        connection_string += f"trustservercertificate={trustservercertificate}&"
        connection_string += f"driver={driver.replace(' ', '+')}"
    else:
        # Иначе используем доверенное подключение
        connection_string = f"mssql+pyodbc://{server}/{database}?"
        connection_string += f"trusted_connection={trusted_connection}&"
        connection_string += f"trustservercertificate={trustservercertificate}&"
        connection_string += f"driver={driver.replace(' ', '+')}"

    return connection_string

if __name__ == "__main__":
    settings_file = 'project/settings.txt'
    settings = read_settings(settings_file)
    connection_string = create_connection_string(settings)
    print(connection_string)
