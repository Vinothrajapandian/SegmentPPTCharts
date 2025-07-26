from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE

COLOR_PALETTE = [
    RGBColor(31,119,180),
    RGBColor(255,127,14),
    RGBColor(44,160,44),
    RGBColor(214,39,40),
    RGBColor(148,103,189),
    RGBColor(140,86,75),
    RGBColor(227,119,194),
    RGBColor(127,127,127),
    RGBColor(188,189,34),
    RGBColor(23,190,207),
]

SUPPORTED_TYPES = {
    XL_CHART_TYPE.BAR_CLUSTERED,
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    XL_CHART_TYPE.LINE,
    XL_CHART_TYPE.AREA
    # You can add more types here
}

def safe_get_values(series):
    try:
        return [v for v in series.values if v is not None]
    except:
        return []

def process_pptx(path_in, seg_count, seg_names):
    prs = Presentation(path_in)
    color_map = COLOR_PALETTE[:seg_count]

    for slide_index, slide in enumerate(prs.slides):
        for shape_index, shape in enumerate(slide.shapes):
            if not shape.has_chart:
                continue

            chart = shape.chart
            chart_type = chart.chart_type

            if chart_type not in SUPPORTED_TYPES:
                continue  # Skip unsupported charts for now

            try:
                series = chart.series
                if not series:
                    continue

                categories = [pt.label for pt in chart.plots[0].categories]
                original_values = safe_get_values(series[0])

                if not categories or not original_values:
                    continue

                # If category count and value count mismatch, skip
                if len(categories) != len(original_values):
                    continue

                new_data = ChartData()
                new_data.categories = categories

                for i in range(seg_count):
                    new_data.add_series(seg_names[i], original_values)

                chart.replace_data(new_data)

                # Update series colors and names
                for idx, s in enumerate(chart.series):
                    s.name = seg_names[idx]
                    fill = s.format.fill
                    fill.solid()
                    fill.fore_color.rgb = color_map[idx % len(color_map)]

            except Exception as e:
                print(f"Error processing slide {slide_index}, shape {shape_index}: {e}")
                continue

    output_path = path_in.replace(".pptx", "_segmented.pptx")
    prs.save(output_path)
    return output_path
