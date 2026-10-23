"""Impress operations (``impress.*``).

helpers.py      slide / shape addressing, layouts, text and styling helpers
read.py         get_overview, get_slide, get_shape, get_text, get_notes, find_text
slides.py       delete / duplicate / move, layout, size, visibility, background, transitions
content.py      add_slide(s), title / body / shape text, speaker notes
shapes.py       text boxes, shapes, lines, images, geometry, order, align, duplicate
tables.py       tables
formatting.py   text formatting, shape styling, hyperlinks
master.py       master slides and built-in themes
"""

from libreoffice_mcp.worker.ops.impress import (  # noqa: F401  (imports register the ops)
    content,
    formatting,
    master,
    read,
    shapes,
    slides,
    tables,
)
