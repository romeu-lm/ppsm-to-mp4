# ppsm-to-mp4
Batch export PowerPoint (.ppsm) files to MP4 using Windows COM automation.

## Requirements
- Windows
- Microsoft PowerPoint installed
- Python 3.9+
- pywin32

## Installation
pip install -r requirements.txt

## Usage
python ppsm_to_mp4.py

## How It Works
The script uses PowerPoint's COM interface via pywin32 to:
1. Open each .ppsm file
2. Call CreateVideo()
3. Wait for export completion
4. Save output to ./Videos
