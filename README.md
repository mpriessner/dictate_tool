# Dictation Tool

A Python-based speech-to-text application that allows users to dictate text using their microphone and automatically paste it where needed.

## Installation

### Prerequisites
- Windows OS
- Python 3.12
- Working microphone
- Internet connection
- OpenAI API key

### Setup Instructions

1. **Create and Activate Conda Environment**
   ```bash
   conda create -n dictate_env python=3.12
   conda activate dictate_env
   ```

2. **Install Required Packages**
   
   First, install PyQt5 through conda:
   ```bash
   conda install -c conda-forge pyqt=5.15.9
   ```

   Then install the remaining requirements:
   ```bash
   pip install openai==0.28.0
   pip install keyboard==0.13.5
   pip install pyperclip==1.8.2
   pip install numpy
   pip install sounddevice
   pip install pywin32==306
   ```

3. **Configure OpenAI API Key**
   - Open `dictate_app.py`
   - Replace `YOUR-API-KEY-HERE` with your OpenAI API key:
     ```python
     openai.api_key = "YOUR-API-KEY-HERE"
     ```

## Usage

1. **Start the Application**
   ```bash
   python dictate_app.py
   ```

2. **Using the Tool**
   - Click and hold the "Press and Hold to Record" button
   - Speak clearly into your microphone
   - Release the button when done speaking
   - Within 2 seconds, click where you want the text to appear
   - The transcribed text will be automatically pasted

## Troubleshooting

- **Audio Issues**: Verify microphone connection and permissions
- **Pasting Problems**: Ensure app has permissions for keyboard input
- **API Errors**: Check OpenAI API key validity and credit balance

## Support

For issues or questions, please open an issue in the repository.
