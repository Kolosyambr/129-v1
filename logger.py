import logging

logging.basicConfig(
    filename="bot_logs.txt",
    level=logging.INFO,
    format="%(asctime)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    encoding="utf-8"  # Указываем кодировку UTF-8 для корректного отображения русских символов
)

def log_message(user_id, username, message_type, text):
    user_info = f"User: {user_id} | Username: {username or 'N/A'}"
    log_text = f"{user_info} | Type: {message_type} | Text: {text}"
    logging.info(log_text)

