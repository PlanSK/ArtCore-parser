import hashlib
import os
import json


def checksum_gen(file_name: str) -> str:
    with open(file_name,'rb') as file:
        return hashlib.md5(file.read()).hexdigest()


def checksum_check(check_file: str, checksum_data: dict, gsheet: bool = False) -> bool:
    _, get_file_name = os.path.split(check_file)
    if (checksum_data.get(get_file_name) and
            checksum_data[get_file_name] == checksum_gen(check_file)):
        return True

    return False


def checksum_list(file: str) -> dict:
    get_checksums = dict()

    try:
        with open(file, 'r', encoding='utf-8') as checksum_data:
            get_checksums = json.load(checksum_data)
    except FileNotFoundError:
        print('Checksum file is not opened.')
    finally:
        return get_checksums


def checksum_dict(get_dict: dict) -> str:
    checksum = hashlib.md5(
        json.dumps(get_dict, sort_keys=True, ensure_ascii=True
    ).encode('utf-8')).hexdigest()

    return checksum