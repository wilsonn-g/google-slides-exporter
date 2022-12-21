import argparse
import re
from typing import List
from time import sleep

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.actions.wheel_input import ScrollOrigin
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys

from PIL import Image
import io  # For Reading the Byte File retrieved from selenium


def create_driver_instance() -> webdriver:
    """
    Creates a chrome webdriver instance.
    :return: Chrome webdriver instance
    """
    chrome_options = Options()
    chrome_options.headless = True
    chrome_driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options
    )
    chrome_driver.maximize_window()
    return chrome_driver


def scrape_title(driver: webdriver) -> str:
    """
    Get the title tag of the HTML document
    :param driver: Selenium webdriver instance
    :return: Text value of the <title> web element
    """
    return driver.title


def get_num_of_slides(driver: webdriver) -> int:
    """
    Get the number of total slides in the presentation
    :param driver: Selenium webdriver instance
    :return: number of slides in the presentation
    """
    try:
        # Text within total slide count element is empty on load until filmstrip is scrolled on
        filmstrip = driver.find_element(By.ID, "filmstrip")
        scroll_origin = ScrollOrigin.from_element(filmstrip, 0, 200)
        ActionChains(driver)\
            .scroll_from_origin(scroll_origin, 0, 400)\
            .perform()

        total_slides = driver.find_element(By.ID, "punch-total-slide-count")
        driver.execute_script("arguments[0].style.display = 'block';", total_slides)  # total_slide element goes to display:none when not scrolling
        return int(re.findall("\d+", total_slides.text)[0])
    except NoSuchElementException:
        print("Error while looking for an element while calculating number of slides")
        return 0
    except IndexError:
        print("Error while performing regex to find total number of slides")
        return 0


def download_slides(slides_url: str, driver: webdriver) -> List:
    """
    Screenshots every slide in the slide deck
    :param driver: Selenium webdriver instance
    :param slides_url: url to the Google Slides presentation
    :return: List of [PIL.Images], each element representing a slide in the presentation
    """
    try:
        driver.get(slides_url)
    except Exception:
        print(f'Ran into exception while trying to download slide {slides_url}')
        return []

    num_slides = get_num_of_slides(driver)
    slides = []
    ActionChains(driver)\
        .key_down(Keys.CONTROL)\
        .send_keys(Keys.F5)\
        .key_up(Keys.CONTROL)\
        .perform()  # Enter presentation mode for high rez images
    sleep(1)  # give webdriver a second to load into full-screen otherwise you will get a black image in pdf

    for i in range(0, num_slides):
        slide_screenshot = screenshot_slide(driver)
        slides.append(slide_screenshot)
        ActionChains(driver)\
            .send_keys(Keys.ARROW_DOWN)\
            .perform()

    ActionChains(driver)\
        .send_keys(Keys.ESCAPE)\
        .perform()
    return slides


def screenshot_slide(driver: webdriver) -> Image:
    """
    Screenshots the specified element
    :param element: HTML element
    :param driver: Selenium webdriver instance
    :return: A screenshot of the slide as a PIL.Image
    """
    image = driver.get_screenshot_as_png()
    buf = io.BytesIO(image)
    return Image.open(buf).convert("RGB")


def sanitize_filename(filename: str) -> str:
    """
    Removes any OS reserved characters in a string
    :param filename: file name to be sanitized
    :return: A valid filename string
    """
    forbidden_chars = '"*\\/\'.|?:<>'
    sanitized_output = ''.join([x if x not in forbidden_chars else '' for x in filename])

    return sanitized_output


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        prog="SlidesExporter",
        usage="main.py [-h] [--url] [SLIDES_URL] [--slides] [FILE]\n Example: python main.py "
              "--url https://docs.google.com/presentation/d/1pA4QO0WEVGbTMpmKBV_1n3458PKxtvvFzDKZi_rsgAo\n "
              "python main.py --slides ./slides-list.txt",
        description="Exports Google Slides presentations as a PDF even if the exporting is disabled by the owner",
    )

    parser.add_argument(
        "--url",
        help="URL link to the Google Slides presentation you want to copy",
        type=str,
    )

    parser.add_argument(
        "--slides",
        help="path to a text file containing all Slides URLs (each have to be on a newline)",
        type=str,
    )

    args = parser.parse_args()

    if args.slides is None and args.url is None:
        parser.error("Supply either a google slides URL or a new-line delimited file with multiple slide urls")

    driver = create_driver_instance()

    presentations_to_download = []

    if args.url:
        presentations_to_download.append(args.url)

    if args.slides:
        try:
            with open(args.slides, "r") as file:
                presentations_to_download = presentations_to_download + file.read().splitlines()
        except OSError:
            print(f"Failed to open file {args.slides}")
            parser.print_help()

    for presentation in presentations_to_download:
        slide_deck = download_slides(presentation, driver)
        title = sanitize_filename(scrape_title(driver))
        if len(slide_deck) > 0:
            slide_deck[0].save(f"{title}.pdf", save_all=True, append_images=slide_deck[1:], quality=95)

    driver.close()
