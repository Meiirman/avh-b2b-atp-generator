def generate(rebder_data) -> dict:
    if rebder_data is None:
        return {"status" : "error", "message" : "Вызов метода не может быть пустым"}
    return {"status" : "success", "message" : "Генерация прошла успешно"}   