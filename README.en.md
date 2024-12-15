# Images to PowerPoint Tool

This is a Python tool that automatically imports PNG images from a folder into a PowerPoint presentation. Images are resized to a 16:9 aspect ratio and sorted by numerical values in their filenames.

## Installation

### Using Pip

```bash
pip install -r requirements.txt
```

### Using Poetry

```bash
poetry install
```

## Usage

Run the tool with the following command:

```bash
python main.py
```

1. Select the input folder containing PNG images.
2. Choose the output file location (default: `output.pptx` in the input folder).
3. Click "Start Processing." The output file will be saved to the specified location upon completion.

## Notes

- Only PNG images are supported.
- Images will be cropped to a 16:9 aspect ratio. Ensure critical content is centered.
- Images are sorted by numerical values in their filenames.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
