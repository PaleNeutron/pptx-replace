# pptx-replace

A python package for replaceing text, images and tables in pptx files

```python
import pptx
import pptx_replace

prs = Presentation("tests/templates/test_template.pptx")
replace_text(prs, "{Main title}", "this is main report title")
slide = prs.slides[1]
replace_text(slide, "{title}", "This is a title")
replace_picture(prs.slides[0], "image.png", auto_reshape=True)
```

## installation

```bash
pip install pptx-replace
```

If you want to put `altair` picture into pptx file, you need install some extra packages.

```bash
pip install "pptx-replace[alt]"
```

see: <https://github.com/altair-viz/altair_saver>

### extra dependency

Depends on your usage, if you want to export table and keep style in jupyter, `selemium` and browser driver is required.

see: <https://github.com/dexplo/dataframe_image/>


## usage

First open your pptx file.

```python
import pptx
import pptx_replace

prs = Presentation("tests/templates/test_template.pptx")
```


### replace text

Replace any text in your ppt

```python
# repalce all occurances of {Main title} in pptx
replace_text(prs, "{Main title}", "this is main report title")
slide = prs.slides[1]

# replace in just one slide
replace_text(slide, "{title}", "This is a title")

replace_text(slide, "{content}", "a quick brown fox jumps over the lazy dog\n" * 5)
```

### replace picture

Replace picture just by matplotlib figuer!

```python
import matplotlib.pyplot as plt

plt.plot([1, 2, 3, 4])
fig = plt.gcf()

replace_picture(prs.slides[1], fig, auto_reshape=False, order="l2r")
```

### replace table

replace table by pandas dataframe

```python
import pandas as pd
import numpy as np

df = pd.DataFrame(np.random.rand(6, 10))
replace_table(slide, df)

```
