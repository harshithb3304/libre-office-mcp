"""Impress tools (``impress_*``): presentations (.odp, .pptx, .ppt).

    read.py        reading: overview, slide, shape, text outline, notes, find, masters
    slides.py      add / delete / duplicate / move slides, layout, size, background, transitions
    content.py     titles, body text, notes, text boxes, shapes, lines, pictures
    tables.py      tables
    arrange.py     move / resize / rotate, stacking order, delete, duplicate, align shapes
    formatting.py  text formatting, shape styling, hyperlinks
    master.py      master slides and built-in themes
    blocks.py      typed slide blocks for impress_add_slides

Each tool forwards to the worker operation of the same name (``impress_x`` -> ``impress.x``).
"""

from libreoffice_mcp.tools.impress import (  # noqa: F401  (imports register the tools)
    arrange,
    content,
    formatting,
    master,
    read,
    slides,
    tables,
)
