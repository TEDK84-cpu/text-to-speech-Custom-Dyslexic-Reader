# Text-to-Speech Application

A portable text-to-speech application with OCR capabilities that can run from a USB stick. This application allows you to select text from your screen and have it read aloud, making it useful for accessibility and productivity purposes.

## Features

- Text-to-speech conversion
- Screen text selection
- OCR (Optical Character Recognition) support
- Portable - runs from USB stick
- Customizable text settings
- Hotkey support

## Prerequisites

1. **Python Installation**
   - Download and install Python 3.8 or later from [python.org](https://www.python.org/downloads/)
   - During installation, make sure to check "Add Python to PATH"
   - Restart your computer after installation

2. **Tesseract-OCR**
   - Download Tesseract-OCR from [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
   - Extract the Tesseract-OCR folder to the same directory as this application

3. **FFmpeg**
   - The setup script will automatically download and install FFmpeg
   - If automatic installation fails, you can manually download FFmpeg from [ffmpeg.org](https://ffmpeg.org/download.html)
   - Place the FFmpeg executable in the `ffmpeg` folder of the application

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/TEDK84-cpu/text-to-speech-Custom-Dyslexic-Reader.git
   cd text-to-speech-Custom-Dyslexic-Reader
   ```

2. Run the setup script:
   ```bash
   setup.bat
   ```

## Folder Structure

After installation, your folder structure should look like this:

```
text-to-speech-Custom-Dyslexic-Reader/
├── .venv/                      # Python virtual environment (created by setup)
├── ffmpeg/                     # FFmpeg executables
│   ├── ffmpeg.exe
│   ├── ffprobe.exe
│   └── ffplay.exe
├── Tesseract-OCR/             # Tesseract OCR files
│   ├── tesseract.exe
│   └── ... (other Tesseract files)
├── Text-to-Speech.py          # Main application file
├── setup.bat                  # Setup script
├── run.bat                    # Run script
├── requirements.txt           # Python dependencies
├── text_settings.json         # Application settings
├── README.md                  # This file
└── LICENSE                    # MIT License file
```

## Usage

1. Run the application:
   ```bash
   run.bat
   ```

2. Use the following hotkeys:
   - `Ctrl+Shift+S`: Start text selection
   - `Ctrl+Shift+R`: Start reading selected text
   - `Ctrl+Shift+X`: Stop reading

## Troubleshooting

If you encounter any issues:

1. **Python Issues**
   - Make sure Python is installed and added to PATH
   - Try running `python --version` in Command Prompt to verify installation
   - If Python is not found, reinstall Python and check "Add Python to PATH"

2. **Audio Issues**
   - Make sure FFmpeg is properly installed (check the `ffmpeg` folder)
   - Verify that your computer's audio is working
   - Try running the application with administrator privileges

3. **OCR Issues**
   - Ensure Tesseract-OCR is properly installed
   - Check that the Tesseract-OCR folder is present in the application directory

4. **General Issues**
   - Try running `setup.bat` again if you get any dependency errors
   - Check that all required files are present in the application folder

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)
- [FFmpeg](https://ffmpeg.org/)
- [Python](https://www.python.org/) 