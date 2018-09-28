# PPTXProgressbar
add a progressbar to pptx slides

为pptx幻灯片添加一个进度条

## 实现
使用python-pptx包画两个矩形再添加文字

## Usage
@see add_pgb.py：
```
from pptx import Presentation
from pgb_adder.pgb_adder import add_pgb

p=Presentation('XXX.pptx')
add_pgb(p)
p.save('XXX2.pptx')
```
