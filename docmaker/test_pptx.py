from pptx import Presentation

prs = Presentation('assets/samplepptx.pptx')

for x in prs.slides:
    print(x.slide_id)
