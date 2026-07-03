# PhyX

If you are only a user, please directly visit the [Woke Dawu Experiment Tool](https://dawu.feixu.site/) website.

## Directory Structure

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

## Packages That Must Be Installed in Advance

- chardet
- Flask
- latex2mathml
- lxml
- matplotlib
- numpy
- openpyxl
- pandas
- python-docx
- scipy
- sympy
- uncertainties

You can install them in batches with `pip`:

```bash
pip install -r requirements.txt
```

Or configure the environment with `conda`:

```bash
conda create -n phyx -c conda-forge python=3.9.6 pip chardet flask lxml matplotlib numpy openpyxl pandas python-docx scipy sympy uncertainties
conda activate phyx
pip install latex2mathml
```

## Development Process for Each Data Processing Program

1. Write the data processing program and place it in the [module](module) folder.
2. Place the experiment guide PDF (the original handout is sufficient), the sample data CSV (the sample data mainly demonstrates the format; the data does not need to be correct or reasonable), and the PNG (method for generating the PNG file: select all and copy in Excel, paste into a QQ chat box or Windows Paint, and then save as; if there is too much data, use ellipses, referring to exp23d; if anything needs special explanation, mark it in red font, referring to exp18) in the static/experiment/expID/ directory. The file name **must** be consistent with the return value of the `name()` function in the data processing program.
3. Modify the module list [modulelist.py](modulelist.py).
4. Modify the front-end code [templates/index.html](templates/index.html).
5. Start [main.py](main.py) locally for test runs.

(Note: the front-end and front-end/back-end connection parts have already been written uniformly, so it is not necessary to write a webpage for each experiment.)

## Development Standards and Notes

- Python 3.9.6 is recommended.
- File name: expID.py, such as exp1.py (if one experiment has multiple sub-experiments, append a/b/... after the ID, such as exp1a.py).
- Functions
  - `name` returns the experiment name.
  - `handle` processes data and generates documents.
- Variable naming convention: when a variable name consists of multiple words, use underscores `_` (such as tight_layout), rather than capitalization (such as tightLayout or TightLayout).
- Code execution order: calculation -> plotting -> inserting results into Word. Do not insert results into Word while calculating.
- For calculations related to a single data value (such as `pi,sqrt,log`), use functions from the `math` library. For calculations related to a set of data (such as `mean,max,min`), first consider member functions (such as `data.max()`), and then `numpy` functions.
- When using a library/function that has not been used before (such as scipy.signal.savgol_filter), its function **must** be explained in comments.
- Submitted code should not have unnecessary output.
- Plotting and linear fitting (refer to [exp5.py](module/exp5.py))
  - Plot with object-oriented plotting (`fig, ax = plt.subplots()`).
  - Set minor ticks to half of major ticks, while keeping major ticks as default.
  - Ticks face inward (already set uniformly in main.py).
  - If a figure has only one set of points and lines, use red for points (`color='r'`) and blue for lines (`color='b'`), and draw the line above the points. If a figure has multiple sets of points and lines, points and lines in the same set should have the same color, using blue (b), red (r), green (g), purple (m), orange (orange), and cyan (c) in order.
  - Use solid circles (`"o"`) for point markers. If a figure has multiple sets of points and lines, use solid circles (o), squares (s), upward triangles (^), diamonds (D), downward triangles (v), and asterisks (*) in order.
  - Use `linewidth=1.5` for line width and `markersize=3` for point size. They may be adjusted appropriately according to the amount of data and the number of data sets, but consistency should be maintained.
  - For plotting dual-y-axis figures, refer to [exp15c.py](module/exp15c.py).
  - Figures with only one set of points and lines generally do not display a legend.
  - Use SourceHanSansSC-Regular.otf as the image font.
  - Axis labels and titles support Latex.
- For uncertainty calculation and linear regression, see the [Data Processing API Guide](数据处理API指南.md).
- Nonlinear fitting can generally be converted into linear fitting through coordinate transformation. If it cannot be converted, use `scipy.optimize.curve_fit` (refer to [exp15c.py](module/exp15c.py)).
- About documents
  - Use Microsoft YaHei as the font.
  - The first line of the document is the experiment name (that is, the return value of the `name()` function), followed by "【Latex 代码在下面，请向下翻阅】".
  - Word formulas for the calculation process are placed in the first half of the document, and Latex formulas are placed in the second half. For the formula insertion API, see the [Formula Insertion API Guide](公式插入API指南.md).
  - Paragraphs with large content spans should be separated by one blank line.
  - Data inserted into documents generally keeps 4 or 5 significant digits (`'%.5g' % x`), and the correlation coefficient $r$ of linear fitting is recommended to keep 8 significant digits.
  - If an image happens to be at the beginning of page 2 while there is a large blank area at the end of page 1, to avoid misunderstanding, add "【本文档不只有一页，请向下翻阅】" after the last paragraph on page 1.
  - For inserting tables, refer to [exp5.py](module/exp5.py).

## Collaboration Method

Collaborating on Github is very simple. You only need to follow the four steps below.

### Create a Fork

There is a `Fork` button in the upper-right corner of this page. Clicking it will show the following interface:

![New fork](https://s2.loli.net/2022/08/15/5FskUI1WhOql3n8.png)

Just click `Create fork`.

### Make Changes

`Create fork` will create a Repository under your account. Its content is the same as the content of this Repository, but you have all permissions. At this point, you can freely make changes in your Repository.

![Commit changes](https://s2.loli.net/2022/08/15/wKltBaYsIj8ASpW.png)

### Prepare a Pull Request

After you have made your changes, you can submit these changes to us to improve this project. On the homepage of your Repository, follow the operation shown in the image below:

![Open pull request](https://s2.loli.net/2022/08/15/TbqXjed3lOhA4Jv.png)

### Submit the Pull Request

In the `Open a pull request` form, submit the content directly to our `main` base. After filling out the form, click `Create pull request`.

![Create pull request](https://s2.loli.net/2022/08/15/4krCp8MSNnehH7T.png)

**At this point, you have successfully submitted your changes.**

Afterward, your submission will appear in the `Pull requests` tab of this project. We will gratefully accept your changes or discuss them further with you.

![Merge pull request](https://s2.loli.net/2022/08/15/s3CrZJvXItwyxgn.png)

Alternatively, if you have good ideas, you are also welcome to [open an issue](https://github.com/feixukeji/PhyX/issues).

> References: [Fork a repo](https://docs.github.com/en/get-started/quickstart/fork-a-repo), [Pull requests](https://docs.github.com/en/pull-requests), [Creating an issue](https://docs.github.com/en/issues/tracking-your-work-with-issues/creating-an-issue#creating-an-issue-from-a-repository)

## Learning References

- Python 3: [Documentation](https://docs.python.org/3/), [Tutorial](https://docs.python.org/3/tutorial/), [Tutorial (CN)](https://www.runoob.com/python3/python3-tutorial.html)
- NumPy: [Documentation](https://numpy.org/doc/), [Learn](https://numpy.org/learn/), [Tutorial (CN)](https://www.runoob.com/numpy/numpy-tutorial.html)
- Matplotlib: [Documentation](https://matplotlib.org/stable/index.html), [Tutorial](https://matplotlib.org/stable/tutorials/index.html), [Tutorial (CN)](https://www.runoob.com/matplotlib/matplotlib-tutorial.html)
- Pandas: [Documentation](https://pandas.pydata.org/docs/), [User Guide](https://pandas.pydata.org/docs/user_guide/index.html), [Tutorial (CN)](https://www.runoob.com/pandas/pandas-tutorial.html)
- Python-docx: [Documentation](https://python-docx.readthedocs.io/en/latest/), [Quickstart](https://python-docx.readthedocs.io/en/latest/user/quickstart.html)

## Contributors

- Organization and planning, front end, front-end/back-end connection program, formula insertion API, data processing program examples, development documentation, and code review: Xulei Sun
- Data processing API: Xulei Sun, Xuehan Zhang, Xuran Zhou, Guanlin Yin
- Technical and security support: Yi Zhao
- Data processing programs corresponding to each experiment:
  |ID|Experiment|Category|Developer|
  |-|-|-|-|
  |0|General tools|General|Xulei Sun|
  |1|Measurement of gravitational acceleration|Mechanics and Thermodynamics|Qin Qin|
  |2|Surface tension|Mechanics and Thermodynamics|Yi Zhao|
  |3|Viscosity coefficient|Mechanics and Thermodynamics|Xulei Sun|
  |4|Measurement of mass and density|Mechanics and Thermodynamics|Zhengting Bao|
  |5|Young's modulus of steel wire|Mechanics and Thermodynamics|Xulei Sun|
  |6|Shear modulus|Mechanics and Thermodynamics|Zhengting Bao|
  |7|Specific heat of solids|Mechanics and Thermodynamics|Xuehan Zhang|
  |8|Uniformly accelerated motion|Mechanics and Thermodynamics|Xulei Sun|
  |9|Speed of sound measurement|Mechanics and Thermodynamics|Xuehan Zhang|
  |10|Magnetic pendulum|Mechanics and Thermodynamics|(To be developed)|
  |11|Semiconductor thermometer|Electromagnetism|Xuehan Zhang|
  |12|Use of an oscilloscope|Electromagnetism|Zhengting Bao|
  |13|Rectification and filtering|Electromagnetism|Yi Zhao|
  |14|DC power supply characteristics|Electromagnetism|Xuehan Zhang|
  |15|Silicon photovoltaic cell|Electromagnetism|Xilin Xia|
  |16|RGB color matching|Electromagnetism|Qin Qin|
  |17|Digital thermometer|Electromagnetism|Zhengting Bao|
  |18|Spectrometer|Optics|Yi Zhao|
  |19|Measurement of small quantities by interferometry|Optics|Yi Zhao|
  |20|Lens parameter measurement|Optics|Xulei Sun|
  |21|Use of a microscope|Optics|Zhengting Bao|
  |22|Diffraction experiment|Optics|Zhengting Bao|
  |23|Photoelectric effect|Modern Physics|Xuehan Zhang|
  |24|Millikan oil-drop experiment|Modern Physics|Qin Qin|
  |25|Physics experiments in daily life|Daily Life|Qin Qin|

## To Do

1. Development of data processing programs for Level 2, Level 3, and Level 4 college physics experiments.
2. Direct input and display on the web side (currently only file upload and download are supported).
3. Handwritten table digit recognition: convert handwritten experimental data into Excel (CSV) files.
4. Online PDF preview (currently unavailable for preview in some mobile browsers).

## License

This project is licensed under the [GNU Affero General Public License v3.0 (AGPL-3.0)](./LICENSE).
