import configparser


def read_config():
    _config = configparser.ConfigParser()
    _config.read('config.ini')
    pepper_mac = _config.get('Addresses', 'pepper_mac', fallback="")
    bt_uuid = _config.get('Bluetooth', 'uuid', fallback="")
    version = _config.get('Generals', 'version', fallback="")

    config_values = {
        'pepper_mac': pepper_mac,
        'bt_uuid': bt_uuid,
        'version': version
    }
    return config_values

