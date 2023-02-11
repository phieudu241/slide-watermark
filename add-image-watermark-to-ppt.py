import aspose.slides as slides
import aspose.pydrawing as drawing

# load presentation
with slides.Presentation("presentation.pptx") as presentation:
    # select slide
    slide = presentation.slides[0]

    # set watermark position
    center = drawing.PointF(presentation.slide_size.size.width / 2, presentation.slide_size.size.height / 2)
    width = 100
    height = 100
    x = center.x - width / 2
    y = center.y - height / 2

    # load image
    with open("watermark.png", "rb") as fs:
        data = fs.read()
        image = presentation.images.add_image(data)

        # add watermark
        watermarkShape = slide.shapes.add_auto_shape(slides.ShapeType.RECTANGLE, x, y, height, width)
        watermarkShape.name = "watermark"

        # set image for watermark
        watermarkShape.fill_format.fill_type = slides.FillType.PICTURE
        watermarkShape.fill_format.picture_fill_format.picture.image = image
        watermarkShape.fill_format.picture_fill_format.picture_fill_mode = slides.PictureFillMode.STRETCH
        watermarkShape.line_format.fill_format.fill_type = slides.FillType.NO_FILL

        # send to back
        slide.shapes.reorder(0, watermarkShape)

        # lock watermark to avoid modification
        watermarkShape.shape_lock.select_locked = True
        watermarkShape.shape_lock.size_locked = True
        watermarkShape.shape_lock.text_locked = True
        watermarkShape.shape_lock.position_locked = True
        watermarkShape.shape_lock.grouping_locked = True

    # save presentation
    presentation.save("image-watermark-ppt.pptx", slides.export.SaveFormat.PPTX)