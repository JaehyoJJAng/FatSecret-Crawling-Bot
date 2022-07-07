import json
import os.path
from pathlib import Path
from typing import Optional


BASE_DIR = Path(__file__).resolve().parent
JSON_FILE = os.path.join(BASE_DIR,'data.json')

def get_word_list(
        key: str,
        default_value: Optional[str] = None,
        json_path : str = JSON_FILE
):

    # JSON 파일 읽어온 후, 변수에 저장
    with open(json_path , 'r' , encoding='UTF-8') as f:
        data = f.read()

    # JSON 파일 dict로 형변환
    data = json.loads(data)

    try:
        return data[key]
    except:
        if default_value :
            return default_value

        raise EnvironmentError(f'Set the {key} environment variable')
