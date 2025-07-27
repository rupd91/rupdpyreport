# rupdpyreport

`rupdpyreport` is a lightweight Python utility for generating PowerPoint and PDF reports from images. It is designed to streamline the process of visual reporting, making it ideal for engineers, analysts, or QA personnel needing structured slide-based documentation of image-based observations.

## Features

- Automatically detects image files (e.g., 01.png, 02.png)
- Supports optional titles and descriptions for each image
- Generates `.pptx` reports
- Converts `.pptx` to `.pdf` using PowerPoint
- Simple API with minimal dependencies

## Installation

To install from GIT using PIP command:

```bash
pip install git+https://github.com/rupd91/rupdpyreport.git


## Usage

### Example 1: Basic image report from current folder

```python
import rupdpyreport

# Auto-detects .png/.jpg images in current directory
pptx_path = rupdpyreport.proc_createpptx()

# Optionally convert to PDF
pdf_path = rupdpyreport.proc_pptx2pdf(pptx_path)
```

### Example 2: Using specific images with custom titles and descriptions

```python
import rupdpyreport

images = ["01.png", "02.png"]
titles = ["Before Repair", "After Repair"]
descriptions = ["Initial condition with damage", "After patch welding"]

pptx_path = rupdpyreport.proc_createpptx(
    images=images,
    titles=titles,
    descriptions=descriptions
)

pdf_path = rupdpyreport.proc_pptx2pdf(pptx_path)
```

## License

This project is licensed under the MIT License. See the LICENSE file for details.

## Author

Rahul Upadhyay
