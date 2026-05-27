# PPTX Parser

A simple Python-based PPTX parser that extracts structured slide data from PowerPoint presentations.

The script parses:

- Text
- Images
- Tables
- Element positions/sizes
- Basic font metadata

Each slide is exported into its own folder containing:

- `meta.json`
- extracted images

This structure is useful for:

- PPT → HTML conversion
- AI ingestion pipelines
- slide reconstruction
- visual editors
- semantic indexing
- coordinate mapping systems

---

# Features

## Extracts

- Text elements
- Images
- Tables
- Shape coordinates
- Element dimensions
- Basic font styles

## Output Structure

```text
output/
└── presentation_name/
    ├── slide_1/
    │   ├── images/
    │   │   ├── image_1.png
    │   │   └── image_2.jpg
    │   │
    │   └── meta.json
    │
    ├── slide_2/
    │   ├── images/
    │   └── meta.json
    │
    └── slide_3/
```

---

# Installation

## 1. Clone Project

```bash
git clone <repo_url>
cd <repo_name>
```

## 2. Install Dependencies

```bash
pip install -r requirements.txt
```

---

# requirements.txt

```txt
python-pptx==1.0.2
Pillow==11.2.1
lxml==5.4.0
```

---

# Usage

Place your `.pptx` file in the project directory.

Update this line inside the script:

```python
PPTX_FILE = "input/sample.pptx"
```

Run:

```bash
python parser.py
```

---

# Output Example

## meta.json

```json
{
    "slide": 1,
    "width": 1280,
    "height": 720,
    "elements": [
        {
            "id": 1,
            "name": "Title 1",
            "shape_type": "TEXT_BOX",
            "x": 80,
            "y": 40,
            "width": 600,
            "height": 80,
            "type": "text",
            "text": "Quarterly Revenue",
            "font": {
                "size": 28,
                "bold": true,
                "italic": false,
                "name": "Calibri"
            }
        },
        {
            "id": 2,
            "name": "Picture 3",
            "shape_type": "PICTURE",
            "x": 700,
            "y": 120,
            "width": 320,
            "height": 240,
            "type": "image",
            "img_url": "./images/image_2.png"
        }
    ]
}
```

---

# Coordinate System

All positions and sizes are exported in pixels.

```json
{
    "x": 100,
    "y": 200,
    "width": 300,
    "height": 150
}
```

These values can directly be used in:

- HTML/CSS absolute layouts
- Canvas rendering
- Fabric.js
- Konva.js
- React-based editors

---

# Supported Components

| Component | Support |
|---|---|
| Text | ✅ |
| Images | ✅ |
| Tables | ✅ |
| Coordinates | ✅ |
| Font Metadata | ✅ |
| Charts | Partial |
| SmartArt | Limited |
| Animations | ❌ |
| Embedded Video | Partial |

PowerPoint is... special. Animations and SmartArt tend to become ancient cursed XML rituals very quickly.


---

# Use Cases

## AI Pipelines

Convert presentations into structured JSON for:

- RAG systems
- embeddings
- semantic search
- LLM ingestion

## PPT Reconstruction

Rebuild slides in:

- HTML
- Canvas
- React
- Fabric.js
- Konva.js

## Migration Systems

Convert PPTs into:

- web editors
- CMS blocks
- design systems
- custom slide platforms

---

# License

MIT
