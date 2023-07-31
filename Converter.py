#imports
import os
import pdfplumber
from pptx import Presentation
from pptx.util import Inches
from pptx.util import Pt
from Summarization import gpt_summarise
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import fitz


# function for converting the pdf to powerpoint
def convert_pdf_to_pptx(pdf_file, output_pptx, pptx_title, pptx_sub_title, slide_theme):
    # create powerpoint
    presentation = Presentation()
    # opening pdf file to extract information from
    with pdfplumber.open(pdf_file) as pdf:
        
        # create title slide including title and subtitle
        slide = presentation.slides.add_slide(presentation.slide_layouts[0])
        left = top = Inches(0)
        pic = slide.shapes.add_picture(slide_theme, left, top, width=presentation.slide_width, height=presentation.slide_height)
        slide.shapes._spTree.remove(pic._element)
        slide.shapes._spTree.insert(2, pic._element)

        title_shape = slide.shapes.title
        title_shape.text = pptx_title

        title_text_frame = title_shape.text_frame
        title_text_frame.paragraphs[0].font.size = Pt(40)
        title_text_frame.paragraphs[0].font.bold = True
        title_shape.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(60, 120, 216)

        subtitle_shape = slide.placeholders[1]
        subtitle_shape.text = pptx_sub_title

        subtitle_text_frame = subtitle_shape.text_frame
        subtitle_text_frame.paragraphs[0].font.size = Pt(18)
        subtitle_text_frame.paragraphs[0].font.bold = False
        subtitle_shape.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(175, 123, 81)
        
        pdf_document = fitz.open(pdf_file)
        # loop through page by page in pdf file
        for page_num in range(len(pdf.pages)):
            # loading information from given page
            page = pdf_document.load_page(page_num)
            # extracting images from page
            image_list = page.get_images(full=True)
            
            for img_index in image_list:
                xref = img_index[0]
                base_image = pdf_document.extract_image(xref)
                image_bytes = base_image["image"]

                image_path = f"temp_img_{page_num}_{xref}.png"
                with open(image_path, 'wb') as img_file:
                    img_file.write(image_bytes)


                slide_image = presentation.slides.add_slide(presentation.slide_layouts[5])

                left = top = Inches(0)
                pic = slide_image.shapes.add_picture(slide_theme, left, top, width=presentation.slide_width, height=presentation.slide_height)
                slide_image.shapes._spTree.remove(pic._element)
                slide_image.shapes._spTree.insert(2, pic._element)

                image_width = Inches(7)
                image_height = Inches(5)
                left = (presentation.slide_width - image_width) // 2
                top = (presentation.slide_height - image_height) // 2

                slide_image.shapes.add_picture(image_path, left, top, width=image_width, height=image_height)
                os.remove(image_path) 
            
            page = pdf.pages[page_num]
            page_text = page.extract_text()
            first_line, *remaining_lines = page_text.split('\n')

            slide = presentation.slides.add_slide(presentation.slide_layouts[5])

            left = top = Inches(0)
            pic = slide.shapes.add_picture(slide_theme, left, top, width=presentation.slide_width, height=presentation.slide_height)
            slide.shapes._spTree.remove(pic._element)
            slide.shapes._spTree.insert(2, pic._element)

            title_shape = slide.shapes.title
            title_shape.text = first_line 

            title_shape.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(60, 120, 216) 

            title_text_frame = title_shape.text_frame
            title_text_frame.paragraphs[0].font.size = Pt(24)
            title_text_frame.paragraphs[0].font.bold = True

            left = Inches(1)
            top = Inches(1.5)
            width = Inches(8)
            height = Inches(4)
            
            textbox = slide.shapes.add_textbox(left, top, width, height)
            text_frame = textbox.text_frame

            text_frame.word_wrap = True

            summarized_content = gpt_summarise('\n'.join(remaining_lines))

            p = text_frame.add_paragraph()
            p.text = summarized_content["text"]
            p.font.size = Pt(18)
            p.font.color.rgb = RGBColor(175, 123, 81)

    presentation.save(output_pptx)
    presentation.open(output_pptx)