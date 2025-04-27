# Dictation and Activity Tracking Tool

A combined tool that provides:
1. Activity tracking to log your work and leisure focus time
2. Voice assistant features for dictation, email composition, and quick Q&A

> **Note**: The dashboard visualization component is currently under development. While the activity tracking works and logs data correctly, the dashboard interface needs work. The core functionality remains usable for tracking activities.

## Installation

### Prerequisites
- macOS or Windows
- Python 3.12
- Working microphone
- Internet connection
- OpenAI API key (for voice assistant features)

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

### Activity Tracking
1. **Start the Focus Timer**
   ```bash
   python focus_timer.py
   ```
   This will track your work and leisure activities, saving logs to the `focus_logs` directory.

### Voice Assistant
1. **Start the Dictation Tool**
   ```bash
   python dictate_app.py
   ```

2. **Using Voice Features**
   - Click and hold the "Press and Hold to Record" button
   - Speak clearly into your microphone
   - Release the button when done speaking
   - Within 2 seconds, click where you want the text to appear
   - The transcribed text will be automatically pasted
   - You can also ask questions or request email composition

## Troubleshooting

- **Audio Issues**: Verify microphone connection and permissions
- **Pasting Problems**: Ensure app has permissions for keyboard input
- **API Errors**: Check OpenAI API key validity and credit balance

## Support

For issues or questions, please open an issue in the repository.
