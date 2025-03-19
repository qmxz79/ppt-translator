# ppt-translator

## Overview
The PPT Translator is a Python application designed to translate PowerPoint presentations from Chinese to English. It utilizes an API for translation and provides a user-friendly graphical interface for ease of use.

## Project Structure
```
ppt-translator
├── src
│   ├── cn2enppt.py       # Main logic for translating PowerPoint presentations
│   ├── gui.py            # Graphical user interface for the application
│   └── utils.py          # Utility functions for file handling and other helpers
├── requirements.txt       # List of dependencies required for the project
└── README.md              # Documentation for the project
```

## Installation
To set up the project, follow these steps:

1. Clone the repository:
   ```
   git clone <repository-url>
   cd ppt-translator
   ```

2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage
1. Run the application:
   ```
   python src/gui.py
   ```

2. Use the interface to select the PowerPoint file you wish to translate.

3. Click the "Translate" button to start the translation process.

4. After translation, use the "Save As" button to choose the location and filename for the translated presentation.

## Dependencies
The project requires the following Python packages:
- `python-pptx`: For handling PowerPoint files.
- `openai`: For accessing the translation API.
- GUI framework (e.g., `tkinter` or `PyQt`): For creating the graphical user interface.

## Contributing
Contributions are welcome! Please feel free to submit a pull request or open an issue for any suggestions or improvements.

## License
This project is licensed under the MIT License. See the LICENSE file for more details.