# Zoom Recorder

A simple application that records your Zoom meetings automatically.

## Prerequisites

- Python 3.7+
- FFmpeg (recommended but not required)
- Zoom client application

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/username/zoom-recorder.git
   cd zoom-recorder
   ```

2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Install FFmpeg (recommended):
   
   FFmpeg is not proprietary software but is recommended for combining audio and video recordings into a single MP4 file.

   **For Windows:**
   - Download from [ffmpeg.org](https://ffmpeg.org/download.html)
   - Add FFmpeg to your system PATH

   **For macOS:**
   ```bash
   brew install ffmpeg
   ```

   **For Linux:**
   ```bash
   sudo apt update
   sudo apt install ffmpeg
   ```

## Zoom Client Configuration

Before using the recorder, configure your Zoom client with these recommended settings:

1. Open Zoom settings
2. Under **Audio**:
   - Enable "Mute my microphone when joining"
3. Under **Video**:
   - Enable "Turn off my video when joining"
4. Under **General**:
   - Enable "Enter full screen automatically when starting or joining a meeting"

## Usage

Run the application with:
```bash
python zoom_recorder.py
```

The application will automatically detect and record your Zoom meetings, saving them as MP4 files in the `recordings` directory.

## Features

- Automatic meeting detection and recording
- Combined audio and video output (requires FFmpeg)
- Minimal CPU usage
- Configurable recording quality

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Issues

If you find a bug or have a feature request, please open an issue on the GitHub repository.
