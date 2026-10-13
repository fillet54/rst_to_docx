========================
reStructuredText to Docx
========================

Render reStructuredText into a Word template with python-docx. Built for
generating documents like a Software Version Description (SVD), where a
fixed Word template wraps generated content: versions, checksums, file lists
and a change log linked to Jira.

.. code-block:: python

   from pathlib import Path
   from rst_to_docx import render

   render(Path("svd.rst"), "template.docx", "svd.docx",
          variables={"release": {"version": "1.4.2"}, "files": [...], "changelog": [...]},
          jira_url="https://jira.example.com/browse/")

``source`` is RST text, or a ``Path`` to an .rst file. See
``examples/build_svd.py`` for a full wrapper and ``examples/svd.rst`` for a
sample document. Run ``python examples/build_svd.py`` to produce
``examples/svd.docx``.

Install: ``pip install -e .``. Requires Python 3.10+, python-docx 0.8.11+ and
docutils 0.18.1+.

The template
============

``examples/make_template.py`` builds a sample template: only the styles
listed below, black Times New Roman text, 12 pt headings numbered 1., 1.1,
1.1.1, and a header and footer on every page (title page included) with the
document number, revision letter and release date, plus the classification
in red at the bottom. It expects ``doc.number``, ``doc.revision``,
``doc.classification``, ``doc.author``, ``release.version``,
``release.date`` and ``project`` in ``variables``, plus the lists
``revisions`` (``rev``, ``date``, ``description``, ``author``) and
``signatures`` (``role``, ``name``) for the revision history and approvals
tables that follow the title page. In practice you design
the template in Word using these markers:

``{{end-instructions}}``
   A paragraph containing only this text ends the instructions page.
   Everything up to and including it (plus a page-break-only paragraph
   right after it) is deleted. Pass ``strip_instructions=False`` to keep it
   while debugging.

``{{content}}``
   A paragraph containing only this text is replaced by the rendered RST.

``{{dotted.name}}``
   Anywhere else in the body, a table, a header or a footer, this is
   replaced from ``variables``. Placeholders split across runs by Word are
   handled. An undefined name raises ``UndefinedVariableError``.

``{{list.*.field}}``
   In a table row, ``*`` stands for each item of a list in ``variables``:
   the row is repeated once per item (and removed if the list is empty).
   Use it for tables that sit outside the RST content, such as a revision
   history or a signature table before ``{{content}}``. Design the header
   row and one data row in Word::

      | Rev                 | Date                 | Description                 |
      | {{revisions.*.rev}} | {{revisions.*.date}} | {{revisions.*.description}} |

   A row can repeat only one list. Other placeholders in the row are filled
   as usual, and cells without placeholders (a blank signature cell) are
   copied as they are.

**Use the instructions page to keep styles alive.** Word only writes a
built-in style into the file once it has been used, and python-docx can only
apply styles that are saved. Put one sample paragraph on the instructions
page for each style the renderer uses:

=====================  ====================================================
RST                    Word style
=====================  ====================================================
section titles         Heading 1 … Heading N (the first level is always Heading 1)
appendix titles        Appendix 1 … Appendix N (custom styles; else Heading N)
paragraphs             Normal
bullet / numbered      List Bullet [2, 3…] / List Number [2, 3…]
2nd+ para in an item   List Continue [2, 3…]
literal / code block   Plain Text (then Macro Text, HTML Preformatted)
table title            Caption
tables, admonitions    Table Grid (table style)
``inline code``        HTML Code (character style)
links                  Hyperlink (character style)
=====================  ====================================================

A missing style raises ``MissingStyleWarning`` and falls back (e.g. Heading 5
→ Heading 4, or direct monospace / blue underline formatting for character
styles).

Heading numbering (1, 1.1, …) belongs in the template: link a multilevel
list to the heading styles. If the template contains fields (e.g. a TOC),
the output asks Word to update them when the file is opened.

RST support
===========

Sections, paragraphs, ``**bold**``, ``*italic*``, ````code````,
``:sub:``/``:sup:``, bullet and numbered lists (nested, with numbering that
restarts for each new list), literal and ``code-block`` blocks, admonitions
(``note``, ``warning``, …, drawn as boxed cells), grid, simple, ``list-table``
and ``csv-table`` tables (with spans, titles and header rows that repeat on
each page), external links, and internal links (``` `Section`_ ``` or
``:ref:``) implemented as Word bookmarks.

Anything else (images, footnotes, definition lists, …) raises
``UnsupportedNodeError``. docutils errors raise ``RstError``, which lists every
problem with its line number. Note that RST does not nest inline markup:
``**:var:`x`**`` is not bold.

Roles and directives
--------------------

``:var:`release.version```
   A value from ``variables``. The dotted path walks dicts, list indices
   and attributes. It must be a scalar.

``:jira:`PROJ-123``` / ``:jira:`text <PROJ-123>```
   A link to ``jira_url + key``, or to ``jira_url`` with ``{key}``
   substituted.

``:ref:`label``` / ``:ref:`text <label>```
   A link to an explicit target (``.. _label:``). The link text defaults to
   the section title.

``:underline:`text```, ``:strike:`text```

``.. appendix::``
   Put this before the first appendix section. Every section after it uses
   the ``Appendix N`` styles instead of ``Heading N``, which the sample
   template letters A., A.1, A.1.1. To do the same in your own template,
   create ``Appendix 1`` … based on the heading styles and link them to a
   multilevel list whose first level is A, B, C.

``.. var-table:: path``
   A table built from a list of dicts (or objects) in ``variables``::

      .. var-table:: changelog
         :columns: key=Issue:jira, type=Type, summary=Summary
         :widths: 15 15 70
         :title: Resolved issues
         :empty: No changes in this release.

   ``key=Header`` pairs set the column order and labels. A trailing ``:jira``
   makes each value (a key, or a list of keys) a Jira link. ``:code`` renders
   the value in monospace.

Extending
---------

Custom roles and directives are ordinary docutils ones. If they emit new
node types, register a renderer:

.. code-block:: python

   from rst_to_docx import register_handler

   @register_handler(MyNode)
   def render_my_node(renderer, node):
       with renderer.inline_paragraph("Normal"):
           with renderer.formatting(bold=True):
               renderer.render_children(node)

Security
========

The RST source and the template are treated as untrusted input:

* ``raw`` is disabled, and so are ``include`` and ``csv-table``
  ``:file:``/``:url:``, which would let a document read local files or fetch
  URLs. Pass ``file_insertion=True`` to ``render`` for RST you trust.
* Links may only use ``http``, ``https`` or ``mailto``. Anything else
  (``file:``, ``javascript:``, UNC paths) raises ``UnsupportedNodeError``.
* XML is parsed by python-docx with entity resolution off, and all text goes
  into the document through the XML API, never by string concatenation.
* ``variables`` lookups never reach attributes whose name starts with ``_``.

The paths passed to ``render`` (source, template, output) are opened as
given. Validating them is the caller's job.

Tests
=====

``pip install -e .[test] && pytest``
