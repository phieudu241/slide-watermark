import aspose.slides as slides
import aspose.pydrawing as drawing

# load presentation
with slides.Presentation("presentation.pptx") as presentation:
    # select slide
    slide = presentation.slides[0]

    # set watermark position
    center = drawing.PointF(presentation.slide_size.size.width / 2, presentation.slide_size.size.height / 2)
    width = 300
    height = 300
    x = center.x - width / 2
    y = center.y - height / 2

    # add watermark
    watermarkShape = slide.shapes.add_auto_shape(slides.ShapeType.RECTANGLE, x, y, height, width)
    watermarkShape.name = "watermark"
    watermarkShape.fill_format.fill_type = slides.FillType.NO_FILL
    watermarkShape.line_format.fill_format.fill_type = slides.FillType.NO_FILL

    # set watermark text, font and color
    watermarkTextFrame = watermarkShape.add_text_frame("Watermark")
    watermarkPortion = watermarkTextFrame.paragraphs[0].portions[0]
    watermarkPortion.portion_format.font_height = 52
    watermarkPortion.portion_format.fill_format.fill_type = slides.FillType.SOLID
    watermarkPortion.portion_format.fill_format.solid_fill_color.color = drawing.Color.red

    # lock watermark to avoid modification
    watermarkShape.shape_lock.select_locked = True
    watermarkShape.shape_lock.size_locked = True
    watermarkShape.shape_lock.text_locked = True
    watermarkShape.shape_lock.position_locked = True
    watermarkShape.shape_lock.grouping_locked = True

    # set rotation
    watermarkShape.rotation = -45

    # send to back
    slide.shapes.reorder(0, watermarkShape)

    # save presentation