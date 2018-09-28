from pptx.util import Length

from pgb_adder.shape_adder import add_rect, add_text


def add_pgb_page(
        slide, cur, total,
        left=0, top=0, height=100000, width_full=12192000,
        fore='89ABBF', past='A0CCEF'
):
    left = Length(left)
    top = Length(top)

    text_height = Length(height * 0.95)
    height = Length(height)

    width_per = Length(width_full * ((cur) / (total)))
    width_full = Length(width_full)

    add_rect(slide, left, top, width_full, height, fore)
    add_rect(slide, left, top, width_per, height, past)
    add_text(slide, left, top, width_full, height, text_height, '%d/%d' % (cur, total))


def add_pgb(
        pptx,
        pgb_left=0, pgb_top=0, pgb_width=12192000, pgb_height=100000,
        pgb_color_fore='89ABBF', pgb_color_past='A0CCEF',
        skip_first=0, skip_last=0, skip_idx=[]
):
    """
    add a progress bar to a pptx
    :param pptx: pptx object from python-pptx @Presentation()
    :param pgb_left: left(Length)
    :param pgb_top: top(Length)
    :param pgb_width: width of the whole progress bar(Length)
    :param pgb_height: height of the whole progress bar, text size is 95% of the progress bar(Length)
    :param pgb_color_fore: color of the right part of the progress bar
    :param pgb_color_past: color of the left part of the progress bar
    :param skip_first: skip the first N slides
    :param skip_last: skip the last N slides
    :param skip_idx: skip the slides whose index is in [N,M,...]
    :return: the pptx object; while python pass the pointer, use the original object is the same
    """
    slides = pptx.slides
    ppt_count = len(slides)

    dismiss = list(range(0, skip_first)) + list(range(ppt_count - 1, ppt_count - skip_last - 1, -1))

    for idx, slide in enumerate(slides):
        if idx in dismiss:
            continue
        if (idx+1) in skip_idx:
            continue

        add_pgb_page(slide, idx + 1, ppt_count,
                     left=pgb_left, top=pgb_top,
                     height=pgb_height, width_full=pgb_width,
                     fore=pgb_color_fore, past=pgb_color_past)

    return pptx