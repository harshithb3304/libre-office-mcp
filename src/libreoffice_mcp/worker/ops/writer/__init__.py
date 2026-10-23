"""Writer operations (``writer.*``).

helpers.py      block addressing, insertion slots, formatting helpers
read.py         get_text, get_structure, find_text, get_tables, ...
content.py      add_paragraph / heading / list / content, edit, delete, replace
formatting.py   format_text, format_paragraph, set_style
tables.py       insert/edit/format tables
page.py         page setup, headers/footers, table of contents
objects.py      images, hyperlinks, bookmarks, footnotes, comments, insert_file
review.py       track changes
"""

from libreoffice_mcp.worker.ops.writer import (  # noqa: F401  (imports register the ops)
    content,
    formatting,
    objects,
    page,
    read,
    review,
    tables,
)
