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

def safe_get_categories(chart):
    try:
        return [pt.label for pt in chart.plots[0].categories]
    except Exception:
        return [f"Category {i+1}" for i in range(len(chart.series[0].values))]

def safe_get_values(series):
    try:
        return [v for v in series.values]
    except Exception:
        return []

def process_pptx(path_in, seg_count, seg_names):
    prs = Presentation(path_in)
    color_map = COLOR_PALETTE[:seg_count]

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_chart:
                continue

            chart = shape.chart

            # Only process charts with at least one series and category
            if not chart.series:
                continue

            try:
                categories = safe_get_categories(chart)
                original_values = safe_get_values(chart.series[0])

                if not original_values or not categories:
                    continue

                # New data with duplicated values for segments
                new_data = ChartData()
                new_data.categories = categories

                for i in range(seg_count):
                    new_data.add_series(seg_names[i], original_values)

                chart.replace_data(new_data)

                # Update series name and color
                for idx, series in enumerate(chart.series):
                    fill = series.format.fill
                    fill.solid()
                    fill.fore_color.rgb = color_map[idx % len(color_map)]
                    series.name = seg_names[idx]

            except Exception as e:
                print(f"Skipping chart due to error: {e}")
                continue  # Move to next chart safely

    output_path = path_in.replace(".pptx", "_segmented.pptx")
    prs.save(output_path)
    return output_path
