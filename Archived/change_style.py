import docx
import docx.shared
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.oxml.ns import qn


def main ():
    
    document = docx.Document()

    minstyle = document.styles.add_style("Min", WD_STYLE_TYPE.PARAGRAPH)
    
    
    for i, styles in enumerate(document.styles):
        print(i, styles)
 
    change_style(minstyle)
    
    text = ("Hello world, this is a new text with a new style and font!")

    document.add_paragraph(text, minstyle)

    document.save("new_paragraph_style_test.docx") 


#style.name = CAP Header
#style.type = PARAGRAPH(1)


""" 
def change_style (style, font_name = "Calibri", 
              font_size = 11, 
              font_bold = True, font_italic = True, font_underline = True, 
              color = RGBColor(0, 0, 0)):

    newfont =  style.font
    newfont.name = font_name
    newfont.size = Pt(font_size)
    newfont.bold = font_bold
    newfont.italic = font_italic
    newfont.underline = font_underline
    newfont.color.rgb = color
    
if __name__ == "__main__" :
    main() 
""" 
def set_new_style (styleobject, font_name = "Calibri", 
              font_size = 11, 
              font_bold = True, font_italic = True, font_underline = True, 
              color = RGBColor(0, 0, 0)):

    arg : 
    output : adds a new paragraph style.type

    newstylename = str(input("Enter new style name : "))
    newstyle = styleobject.add_style(newstylename, WD_STYLE_TYPE.PARAGRAPH)
    newfont =  newstyle.font
    newfont.name = font_name
    newfont.size = Pt(font_size)
    newfont.bold = font_bold
    newfont.italic = font_italic
    newfont.underline = font_underline
    newfont.color.rgb = color
"""


