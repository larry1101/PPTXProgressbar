from pptx import Presentation

from pgb_adder.pgb_adder import add_pgb

# Open a pptx file
p = Presentation('结香厌世00.pptx')

WIDTH_VAL = 12192000
HEIGHT_VAL = 130000

COL_FORE = '89BBAF'
COL_PAST = 'A0DCCF'

add_pgb(p,
        pgb_left=0, pgb_top=0, pgb_width=WIDTH_VAL, pgb_height=HEIGHT_VAL,
        pgb_color_fore=COL_FORE, pgb_color_past=COL_PAST,
        skip_first=1, skip_last=3,
        # skip_idx=[5,8,9]
        )

p.save('结香厌世.pptx')
