"""
Slide 10: ML Price Prediction — Ridge Regression model.
"""

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from theme import (
    BG_COLOR,
    CARD_COLOR,
    ACCENT_INDIGO,
    ACCENT_VIOLET,
    ACCENT_EMERALD,
    ACCENT_AMBER,
    TEXT_DARK,
    TEXT_MEDIUM,
    TEXT_LIGHT,
    WHITE,
    FONT_TITLE,
    FONT_BODY,
    SLIDE_WIDTH,
    SLIDE_HEIGHT,
    CARD_LEFT,
    CARD_TOP,
    CARD_WIDTH,
    CARD_HEIGHT,
    set_bg_solid,
    add_rectangle,
    add_textbox,
    add_pill,
    ACCENT_CORAL,
    PILL_BLUE,
    PILL_MINT,
    PILL_CORAL,
    PILL_SLATE,
)


def _model_step(slide, left, top, w, h, label, fill_color, text_color=WHITE):
    """One step in the ML model pipeline."""
    add_rectangle(slide, left, top, w, h, fill_color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_textbox(
        slide,
        left + Inches(0.12),
        top + Inches(0.10),
        w - Inches(0.24),
        Inches(0.35),
        text=label,
        font_size=11.5,
        font_color=text_color,
        bold=True,
        align=PP_ALIGN.CENTER,
    )


def add_slide(prs, assets_dir):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg_solid(slide, BG_COLOR)

    # White card
    add_rectangle(
        slide,
        CARD_LEFT,
        CARD_TOP,
        CARD_WIDTH,
        CARD_HEIGHT,
        CARD_COLOR,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE,
    )

    # Title
    add_textbox(
        slide,
        Inches(0.55),
        Inches(0.35),
        Inches(12.0),
        Inches(0.6),
        text="ML: Hourly Fee Price Prediction",
        font_size=26,
        font_color=TEXT_DARK,
        bold=True,
        font_name=FONT_TITLE,
    )

    # Subtitle
    add_textbox(
        slide,
        Inches(0.55),
        Inches(0.95),
        Inches(12.0),
        Inches(0.4),
        text="Ridge Regression (L2-regularized) model trained on outlier-free tutor data",
        font_size=13,
        font_color=TEXT_MEDIUM,
    )

    # Model pipeline flow
    steps = [
        (
            "Input Features",
            "Experience, city tier, qualification, occupation, subject, level, mode",
            ACCENT_INDIGO,
        ),
        (
            "Ridge Regression",
            "L2-regularized linear model on design matrix",
            ACCENT_VIOLET,
        ),
        (
            "Prediction",
            "Estimated hourly fee (₹/hr) with confidence band",
            ACCENT_EMERALD,
        ),
    ]
    step_w = Inches(3.9)
    step_h = Inches(1.35)
    gap = Inches(0.2)
    start_x = Inches(0.66)
    y = Inches(1.75)
    for i, (label, desc, color) in enumerate(steps):
        left = start_x + (i * (step_w + gap))
        _model_step(slide, left, y, step_w, step_h, label, color)
        add_textbox(
            slide,
            left + Inches(0.12),
            y + Inches(0.52),
            step_w - Inches(0.24),
            Inches(0.75),
            text=desc,
            font_size=9.5,
            font_color=TEXT_MEDIUM,
            align=PP_ALIGN.CENTER,
        )
        if i < len(steps) - 1:
            # arrow
            add_rectangle(
                slide,
                left + step_w + Inches(0.02),
                y + Inches(0.5),
                Inches(0.28),
                Inches(0.35),
                ACCENT_INDIGO,
                shape=MSO_SHAPE.RIGHT_ARROW,
            )

    # Model metrics (2x2 mini cards)
    metrics = [
        ("R² Score", "Variance explained", ACCENT_INDIGO),
        ("MAE", "Mean Absolute Error", ACCENT_VIOLET),
        ("RMSE", "Root Mean Squared Error", ACCENT_EMERALD),
        ("Records", "Training samples", ACCENT_AMBER),
    ]
    mw = Inches(3.0)
    mh = Inches(1.15)
    my = Inches(3.55)
    mx = Inches(0.66)
    for i, (label, desc, color) in enumerate(metrics):
        col = i % 2
        row = i // 2
        l = mx + (col * (mw + Inches(0.25)))
        t = my + (row * (mh + Inches(0.2)))
        add_rectangle(
            slide,
            l,
            t,
            mw,
            mh,
            PILL_SLATE if hasattr(__import__("theme"), "PILL_SLATE") else CARD_COLOR,
            shape=MSO_SHAPE.ROUNDED_RECTANGLE,
        )
        add_rectangle(
            slide, l, t, Inches(0.07), mh, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE
        )
        add_textbox(
            slide,
            l + Inches(0.22),
            t + Inches(0.10),
            mw - Inches(0.35),
            Inches(0.35),
            text=label,
            font_size=12,
            font_color=TEXT_DARK,
            bold=True,
        )
        add_textbox(
            slide,
            l + Inches(0.22),
            t + Inches(0.48),
            mw - Inches(0.35),
            Inches(0.3),
            text=desc,
            font_size=9.5,
            font_color=TEXT_MEDIUM,
        )

    # Bottom insight
    add_rectangle(
        slide,
        Inches(0.66),
        Inches(5.45),
        Inches(12.0),
        Inches(0.85),
        PILL_BLUE if hasattr(__import__("theme"), "PILL_BLUE") else CARD_COLOR,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE,
    )
    add_textbox(
        slide,
        Inches(0.95),
        Inches(5.58),
        Inches(11.5),
        Inches(0.65),
        text="🤖  The ML model helps tutors price competitively — factoring in experience, location, qualifications, subject, and teaching mode.",
        font_size=12.5,
        font_color=TEXT_MEDIUM,
        font_name=FONT_BODY,
    )

    # Footer
    add_textbox(
        slide,
        Inches(0.78),
        Inches(7.00),
        Inches(5.0),
        Inches(0.25),
        text="BMP Tutors Presentation  •  Slide 10 of 14",
        font_size=10,
        font_color=TEXT_LIGHT,
    )

    return slide
