# Third-party notices

The standalone bundle contains CPython (Python Software Foundation license) and public packages required by pdfplumber and openpyxl. It does not include Codex, Node.js, Microsoft Excel, or a proprietary spreadsheet library.

Package versions are recorded in `components.json` in each installed version. Original license files are retained in `runtime/LICENSE.txt`, each distribution's `.dist-info` directory under `runtime/Lib/site-packages`, and package-specific license directories, including pypdfium2's third-party notices. Review those licenses when redistributing.

Direct packages: pdfplumber 0.11.9 (MIT), openpyxl 3.1.5 (MIT). Dependencies include pdfminer.six, Pillow, pypdfium2/PDFium, cryptography, cffi, pycparser, charset-normalizer and et-xmlfile. Their original distributions and notices are the authoritative license texts.

Windows PowerShell, .NET Framework and Windows components are provided by the user's Windows installation. They are not copied into this bundle. Microsoft's Visual C++ runtime DLLs supplied with the CPython distribution are included for that Python runtime.
