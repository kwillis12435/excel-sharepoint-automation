def log_message(message):
    print(f"[LOG] {message}")

def format_data(data):
    # Assuming data is a list of dictionaries
    formatted_data = []
    for item in data:
        formatted_data.append({k: str(v) for k, v in item.items()})
    return formatted_data