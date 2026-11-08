DEFAULT_MAPPING = {"ص":"M921", "ع":"M1400", "م":"SHIFT", "ن":"OFF", "خالی":"OFF", "ض":"ERROR", "غ":"ERROR"}

def translate(value, mapping):
    key = (value or "").strip() or "خالی"
    return mapping.get(key, "ERROR")
