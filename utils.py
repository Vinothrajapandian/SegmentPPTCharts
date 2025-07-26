from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.dml.color import RGBColor

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

def process_pptx(path_in, seg_count, seg_names):
    prs = Presentation(path_in)
    color_map = COLOR_PALETTE[:seg_count]

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_chart:
                continue
            chart = shape.chart
            if chart.chart_type != XL_CHART_TYPE.BAR_CLUSTERED:
                continue

            data = ChartData()
            categories = [c.label for c in chart.plots[0].categories]
            data.categories = categories

            original_values = [pt.value for pt in chart.plots[0].series[0].values]
            for idx in range(seg_count):
                series_name = seg_names[idx]
                # Example: we repeat original values or can sketch dummy segmentation
                data.add_series(series_name, original_values)

            chart.replace_data(data)
            for idx, series in enumerate(chart.series):
                fill = series.format.fill
                fill.solid()
                fill.fore_color.rgb = color_map[idx]
                series.name = seg_names[idx]

    out_path = path_in.replace(".pptx", "_segmented.pptx")
    prs.save(out_path)
    return out_path
