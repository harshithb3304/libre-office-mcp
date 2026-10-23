"""Writer tools (``writer_*``): word-processing documents (.odt, .docx, .doc, .rtf, .txt).

    read.py      reading: text, structure, outline, find, tables, images, comments...
    content.py   adding / changing / deleting content and formatting text
    tables.py    table creation and editing
    objects.py   images, links, bookmarks, notes, comments, page layout, change tracking
    blocks.py    typed content blocks for writer_add_content

Each tool forwards to the worker operation of the same name (``writer_x`` -> ``writer.x``).
"""

from libreoffice_mcp.tools.writer import (  # noqa: F401  (imports register the tools)
    content,
    objects,
    read,
    tables,
)
