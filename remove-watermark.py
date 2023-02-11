import aspose.slides as slides
import aspose.pydrawing as drawing

# load presentation
with slides.Presentation("text-watermark-slide.pptx") as presentation:
    # select slide
    slide = presentation.slides[0]

    shapesToRemove=[]

    # loop through all the shapes in slide
    for i in range(len(slide.shapes)):
        shape = slide.shapes[i]

        # if shape is watermark
        if shape.name == "watermark":
            shapesToRemove.append(shape)

    # loop through all the shapes to be removed
    for i in range(len(shapesToRemove)):
        # remove shape
        slide.shapes.remove(shapesToRemove[i])

    # save presentation
    presentation.save("remove-watermark.pptx", slides.export.SaveFormat.PPTX)