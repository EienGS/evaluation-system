import json

def load_config():

    with open("config_params.json","r",encoding="utf-8") as f:
        return json.load(f)