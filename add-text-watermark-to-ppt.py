import os.path

import aspose.slides as slides
import aspose.pydrawing as drawing
from os import walk


# Instantiates the License class
license = slides.License()

# Sets the license file path
license.set_license("Aspose.Slides.Product.Family.lic")

slides_folder_path = "slides"
watermark_folder_path = "watermark"
f = []
for (dir_path, dir_names, filenames) in walk(slides_folder_path):
    print(dir_path)
    print(dir_names)
    print(filenames)

    for filename in filenames:
        file_path = os.path.join(dir_path, filename)
        # load presentation
        with slides.Presentation(file_path) as presentation:
            # select slide
            # master = presentation.masters[0]

            for slide in presentation.slides:
                # set watermark position
                center = drawing.PointF(presentation.slide_size.size.width / 2, presentation.slide_size.size.height / 2)
                width = 500
                height = 500
                x = center.x - width / 2 + 50
                y = center.y - height / 2

                # add watermark
                watermarkShape = slide.shapes.add_auto_shape(slides.ShapeType.RECTANGLE, x, y, height, width)
                watermarkShape.name = "watermark"
                watermarkShape.fill_format.fill_type = slides.FillType.NO_FILL
                watermarkShape.line_format.fill_format.fill_type = slides.FillType.NO_FILL

                # set watermark text, font and color
                watermarkTextFrame = watermarkShape.add_text_frame("ngocsensei.com")
                watermarkPortion = watermarkTextFrame.paragraphs[0].portions[0]
                watermarkPortion.portion_format.font_height = 48
                watermarkPortion.portion_format.fill_format.fill_type = slides.FillType.SOLID
                # HEX:  # c0c0c0
                watermarkPortion.portion_format.fill_format.solid_fill_color.color = drawing.Color.from_argb(196, 192, 192, 192)

                # lock watermark to avoid modification
                watermarkShape.shape_lock.select_locked = True
                watermarkShape.shape_lock.size_locked = True
                watermarkShape.shape_lock.text_locked = True
                watermarkShape.shape_lock.position_locked = True
                watermarkShape.shape_lock.grouping_locked = True

                # send to back
                slide.shapes.reorder(2, watermarkShape)

                # set rotation
                watermarkShape.rotation = -25

            # save presentation
            save_file_path = os.path.join(watermark_folder_path, filename)
            presentation.save(save_file_path, slides.export.SaveFormat.PPTX)
