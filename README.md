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
Run the script from the directory containing your `.ppsm` files:
python ppsm_to_mp4.py

The script scans the current working directory for `.ppsm` files and writes the exported videos to:

- `./Videos/*.mp4`

## How It Works
The script uses PowerPoint's COM interface via pywin32 to:
1. Open each `.ppsm` file in the input folder
2. Call `CreateVideo()`
3. Wait for export completion
4. Save output to `./Videos`
