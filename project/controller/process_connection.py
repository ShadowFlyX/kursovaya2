import pyodbc

def get_available_drivers():
    drivers = pyodbc.drivers()

    return list(filter(lambda x: x.startswith("ODBC"), drivers))

def get_latest_sql_server_driver():
    drivers = get_available_drivers()
    sql_server_drivers = [driver for driver in drivers if 'SQL Server' in driver]
    if sql_server_drivers:
        return sorted(sql_server_drivers, reverse=True)[0]
    else:
        raise Exception("No SQL Server ODBC driver found.")

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
    driver = settings.get('DRIVER', get_latest_sql_server_driver())
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
