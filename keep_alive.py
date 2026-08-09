import time
from playwright.sync_api import sync_playwright

APP_URL = "https://social-analysis-chnu25.streamlit.app"

def wake_up_streamlit():
    print(f"🌐 Відкриваємо сайт: {APP_URL}")
    
    with sync_playwright() as p:
        # Запускаємо безголовий Chrome
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        # Переходимо на сайт і чекаємо первинного завантаження
        page.goto(APP_URL, wait_until="networkidle", timeout=60000)
        time.sleep(5)
        
        # Перевіряємо, чи є на сторінці кнопка пробудження
        # Streamlit використовує селектори для кнопок (за текстом або за атрибутом)
        wake_button = page.get_by_role("button", name="Yes, get this app back up!")
        
        if wake_button.is_visible():
            print("😴 Додаток спав! Натискаємо кнопку 'Yes, get this app back up!'...")
            wake_button.click()
            
            # Чекаємо 30 секунд, поки пройде стадія "Your app is in the oven" і сайт розігріється
            print("⏳ Чекаємо розігріву (in the oven)...")
            time.sleep(30)
            print("🎉 Додаток успішно розбуджено!")
        else:
            print("⚡ Додаток вже був активним і готовим до роботи!")
            
        # Робимо додаткову паузу для фіксації активного WebSocket-з'єднання
        time.sleep(5)
        browser.close()

if __name__ == "__main__":
    wake_up_streamlit()
