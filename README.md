# google-slides-exporter

A tool to manually export a Google Slides Presentation as a PDF.

Google Slides already has an export feature, however, document owners are allowed to disable exporting and this is simply a quick workaround.

## Run

- Install dependency packages

      pip install validators webdriver-manager selenium pillow

- OR if you are having trouble with newer versions of the packages

      pip install -r requirements.txt

## Usage

- usage:

      main.py [-h] [--slides] [FILE/URL]

- Example

      python main.py --slides https://docs.google.com/presentation/d/1pA4QO0WEVGbTMpmKBV_1n3458PKxtvvFzDKZi_rsgAo
