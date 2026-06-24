# PPTX Translator

A command-line tool to automatically translate PowerPoint (`.pptx`) presentations while preserving their original formatting, colors, and layout. 

Under the hood, this tool uses `python-pptx` to safely modify slide text and the `translators` package, which leverages free translation engines like Bing, Google, and Alibaba to perform robust string translations.

## Features
- **Format Preservation**: Keeps your original fonts, colors, positioning, and styling intact!
- **Zero API Keys Required**: Uses the `translators` package to query popular free web translation engines.
- **Smart Translation**: Automatically skips translating numbers and pure symbols.
- **Progress Bar**: Shows a beautiful progress bar while translating text blocks.
- **Caching**: Caches unique strings to minimize redundant translations.

## Installation

1. Clone or download this repository.
2. Ensure you have Python 3.7+ installed.
3. Install the dependencies using pip:

```bash
pip install -r requirements.txt
```

## Usage

### Method 1: The Desktop App (Windows Only)
For the easiest experience on Windows, you can use the included `run_translator.bat` script:
1. Double-click `run_translator.bat`. A window will open asking you to paste the path to your `.pptx` file.
2. **Alternatively**, you can drag and drop your `.pptx` file directly onto the `run_translator.bat` file to automatically start the translation!

*(Tip: You can create a desktop shortcut to `run_translator.bat` so you can use it like a regular app from your desktop!)*

### Method 2: Command Line Usage

You can also run the script directly from the terminal for more advanced options:

```bash
python pptx_translator.py --input <path_to_your_presentation.pptx>
```

### Command Line Arguments

- `-i`, `--input`: **(Required)** Path to the input `.pptx` file.
- `-o`, `--output`: *(Optional)* Path to save the translated `.pptx` file. If omitted, the tool will append the target language to the original file name (e.g., `presentation_en.pptx`).
- `-f`, `--from_lang`: *(Optional)* Source language code. Defaults to `"auto"`.
- `-t`, `--to_lang`: *(Optional)* Target language code. Defaults to `"en"`.
- `-e`, `--engine`: *(Optional)* Translation engine to use. Defaults to `"bing"`. Other popular options include `"google"`, `"alibaba"`, `"baidu"`.

### Examples

**1. Basic English Translation (Auto-detects source language)**
```bash
python pptx_translator.py -i my_slides.pptx
```

**2. Translate to Spanish using Google Translate**
```bash
python pptx_translator.py -i my_slides.pptx -t es -e google
```

**3. Explicitly translate from Chinese to English with a custom output name**
```bash
python pptx_translator.py -i my_slides.pptx -o translated_slides.pptx -f zh -t en
```

## Limitations
- This tool does not translate text embedded inside images.
- Free translation engines may impose rate limits. The tool includes a small delay between requests and a retry mechanism to mitigate this.

## Contributing
Feel free to open issues or submit pull requests for additional features and bug fixes!
