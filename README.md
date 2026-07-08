# PhyX

PhyX is a web-based laboratory data processing tool for university physics experiments. It provides experiment-specific workflows for uploading CSV or Excel data, running calculations, generating plots, and exporting reports. The application also includes experiment guides, sample data, uncertainty reference tables, and an in-browser PDF viewer to support physics lab preparation and reporting.

This project is part of Woke 365, a Gold Award winner at USTC's Yuqing Cup Campus Software Design Competition.

## Tech Stack

- **Backend:** Python and Flask
- **Scientific computing:** NumPy, SciPy, pandas, SymPy, and uncertainties
- **Plotting:** Matplotlib with Source Han Sans font support
- **Document generation:** python-docx, latex2mathml, lxml, and MathML-to-OMML conversion
- **Spreadsheet handling:** openpyxl, CSV, XLS, and XLSX uploads
- **Frontend:** HTML templates, JavaScript, Layui, and static assets
- **PDF preview:** Mozilla PDF.js

## Project Structure

```text
│ head.py (universal header for experiment data processing programs)
│ main.py (main program)
│ modulelist.py (module list)
│ requirements.txt (packages that must be installed in advance)
│ SourceHanSansSC-Regular.otf (font file for plotting)
│
├─api (API code)
│      calc.py (data processing API)
│      insert.py (formula insertion API)
│      transformer.py (expression conversion)
│      MML2OMML.XSL (convert to Word objects)
│
├─module (stores data processing programs for each experiment)
│
├─static
│ ├─experiment (stores static files for each experiment)
│ ├─layui (front-end framework)
│ ├─pdf.js (from Mozilla, used for online preview of PDF files)
│ └─uncertainty (tables related to uncertainty)
│
├─templates (front-end code)
│      404.html (error page)
│      experiment.html (specific experiment processing page)
│      index.html (website homepage)
│      uncertainty.html (tables related to uncertainty)
│      viewer.html (online PDF preview page)
│
└─usrdata (stores files generated while the program is running)
```

## Quick Start

1. Create and activate a virtual environment:

   ```bash
   python -m venv .venv
   source .venv/bin/activate
   ```

2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Start the application:

   ```bash
   python main.py
   ```

4. Open `http://127.0.0.1:5000/` in your browser.

For the development guide, see [开发指南.md](./开发指南.md).

## Contributors

Xulei Sun, Xuehan Zhang, Yi Zhao, Zhengting Bao, Qin Qin, Xuran Zhou, Guanlin Yin, Xilin Xia

## License

This project is licensed under the [GNU Affero General Public License v3.0 (AGPL-3.0)](./LICENSE).
