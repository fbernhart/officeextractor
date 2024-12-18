#############
Release Notes
#############

..
    This is a template: Please copy it and then remove indentation!

    X.X.X
    ====================

    Released: YYYY-MM-DD

    **🎉 New**

    * Note: for new, great features
    *

    **💪🏼 Improvements**

    * Note: for smaller improvements
    *

    **🐛 Bug-Fixes**

    * Note: Please reference GitHub issues with `#999 <https://github.com/fbernhart/officeextractor/issues/999>`_ and pull requests with :pr:´999´
    * Please reference GitHub pull requests with `#999 <https://github.com/fbernhart/officeextractor/pull/999>`_

    **⚠️ Deprecation**

    * Note: For any dropped Python versions and dependencies or deprecated features and parameters etc.
    *

    **📘 Documentation**

    *
    *

    **🧹 Cleanup**

    *
    *

    | Thanks to the following people on GitHub for contributing to this release:
    | *GitHub-Name-1*, *GitHub-Name-2* and *GitHub-Name-3* (Note: mention all the merged pull requests since last release here!)

    --------------------------------------------


    This is for the upcoming release. Please fill in the changes of your Pull Request:

0.1.2
====================

Released: 2024-12-03

**🎉 New**

* Added support for Python 3.11, 3.12 and 3.13.


**⚠️ Deprecation**

* Removed support for Python 3.6, 3.7 and 3.8.


0.1.1
====================

Released: 2020-12-01

**🐛 Bug-Fixes**

* Fixed an issue with paths under Windows. ``r"\"`` and ``"\\"`` are now working as expected.
* The output directory can now be the parent directory of the source file. E.g. ``officeextractor.extract(src="some/folder/file.docx", dest="some/folder")``. The output subdirectories are now created with ``Extract_`` as prefix. E.g. ``some/folder/Extracted_file.docx``


0.1.0
====================

Released: 2020-11-17

**🎉 New**

* First release of officeextractor
