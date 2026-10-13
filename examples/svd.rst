.. Software Version Description body. The cover page, TOC, headers and
   footers live in the Word template; this file is everything after them.

Scope
=====

This Software Version Description (SVD) identifies version
:var:`release.version` of :var:`project`, released on
:var:`release.date`. It lists the delivered files with their checksums and
the changes made since version :var:`release.previous`.

See `Delivered Files`_ for the release contents and :ref:`changes` for the
change log.

Referenced Documents
--------------------

.. list-table::
   :header-rows: 1
   :widths: 25 75

   * - Document
     - Title
   * - :var:`doc.srs`
     - Software Requirements Specification
   * - :var:`doc.sdd`
     - Software Design Description

Version Description
===================

Delivered Files
---------------

The files below make up the delivery. Verify each one with
``sha256sum -c SHA256SUMS`` before installation.

.. var-table:: files
   :columns: name=File Name:code, version=Version, size=Size (bytes), sha256=SHA-256:code
   :widths: 30 12 13 45
   :title: Delivered files

Build Environment
-----------------

========== ==============================
Item       Value
========== ==============================
Compiler   :var:`build.compiler`
Commit     :var:`build.commit`
Pipeline   :var:`build.id` (https://ci.example.com/)
========== ==============================

.. _changes:

Changes Since :var:`release.previous`
-------------------------------------

.. var-table:: changelog
   :columns: key=Issue:jira, type=Type, summary=Summary
   :widths: 15 15 70
   :title: Resolved issues
   :empty: No changes in this release.

Known Issues
------------

* :jira:`SVD-140` -- the progress bar can stall at 99% on slow disks.
  Workaround: wait, it finishes.
* Installer messages are English only.

  - Localization is planned for :jira:`release 2.0 <SVD-200>`.

Installation
============

#. Copy the delivered files to the target.
#. Verify checksums::

      sha256sum -c SHA256SUMS

#. Run the installer as an administrator.

.. note::

   Uninstall versions earlier than *1.0* before installing this release.

.. warning::

   Do **not** interrupt the installer while it updates firmware.
