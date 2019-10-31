from wand.image import Image as wi
pdf = wi(filename=u'C:\\Users\\Administrator\\Desktop\\610725105202JC00006.pdf', resolution=300)
pdfImage=pdf.convert("jpeg")
for img in pdfImage.sequence:
    page=wi(image=img)
    page.save(filename=u'C:\\Users\\Administrator\\Desktop\\out.jpg')

