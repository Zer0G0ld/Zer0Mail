import logging
from menu import main_menu
from dotenv import load_dotenv

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
load_dotenv()

if __name__ == __main__:
    main_menu()
