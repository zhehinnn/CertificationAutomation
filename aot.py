import os
from PIL import Image, ImageDraw, ImageFont
import smtplib
from email.message import EmailMessage
import csv
import pandas as pd
import sys
import pathlib

class Message():
    def __init__(self):
        self.RESOURCE_ROOT = None
        self.jpg_path = None
        self.gmailFrom = None
        self.gmailSubject = None
        self.body = None
        self.app_password = None
        # This defines your text font and color
        self.font_path = None
        self.trueType = None
        self.fontColor = None # in RGB
        # This defines how ur text will be placed
        self.anchor = None # Refer to https://pillow.readthedocs.io/en/stable/handbook/text-anchors.html#examples
        self.textPosition = None # in px
        self.textAlign = None # Left, Center, Right
    
    def setTextSettings(self, textPos, anchorVar, fontColor, textAlign):
        self.anchor = anchorVar
        self.textPosition = textPos
        self.textAlign = textAlign
        self.fontColor = fontColor
    
    def setJPGPath(self, path):
        self.jpg_path = os.path.join(self.RESOURCE_ROOT, path)
    
    def sendMessage(self, name, recipient_email):
        output_path = os.path.join(self.RESOURCE_ROOT, "annotated-pdf.jpg")
        pdf_path = os.path.join(self.RESOURCE_ROOT, "NexTech_Participation_Certificate.pdf")
        
        # Opens Certificate template
        img = Image.open(self.jpg_path)
        
        # Adds text to template
        draw = ImageDraw.Draw(img)
        # draw.text((1000, 650), name, fill =(255, 255, 255), align="center", font = self.trueType, anchor = "mb") 
        draw.text(self.textPosition, name, fill = self.fontColor, align = self.textAlign , font = self.trueType, anchor = self.anchor)  ### Refer to https://pillow.readthedocs.io/en/stable/handbook/text-anchors.html#examples ###
        # Saves edited JPG and converts to PDF
        img.save(output_path)
        pdfcon = Image.open(output_path)
        pdf_1 = pdfcon.convert('RGB')
        pdf_1.save(pdf_path)
        
        msg = EmailMessage()
        
        # EDIT THE SUBJECT #
        msg['Subject'] = 'Certification on attending the NexTech Conference & Expo 2024'
        #        #         #
        
        msg['From'] = self.gmailFrom
        msg['To'] = recipient_email
        msg.preamble = 'You will not see this in a MIME-aware mail reader.\n'
        
        # EDIT THIS BODY #
        txtbody = f"""
    Dear {name},

    Greetings from Agents of Tech!

    Thank you for joining us at NexTech Conference & Expo 2024! We hope that the knowledge gained and connections made during NexTech Conference & Expo 2024 will continue to inspire and empower you in your endeavours.

    You may find your e-certificate awarded for your participation attached to this email. Please let us know immediately if there is any error with the certificates. 

    Once again, thank you for your support. We look forward to welcoming you back in our future events!

    Follow us on our IG @agentsoftech.tlc for future exciting events and updates!

        """
        #        #         #
        
        altbody = """
        <html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 15">
<meta name=Originator content="Microsoft Word 15">
<link rel=File-List href="Dear_files/filelist.xml">
<link rel=Edit-Time-Data href="Dear_files/editdata.mso">
<!--[if !mso]>
<style>
v\:* {behavior:url(#default#VML);}
o\:* {behavior:url(#default#VML);}
w\:* {behavior:url(#default#VML);}
.shape {behavior:url(#default#VML);}
</style>
<![endif]--><!--[if gte mso 9]><xml>
 <o:DocumentProperties>
  <o:Author>Zhe Hin _</o:Author>
  <o:LastAuthor>Zhe Hin _</o:LastAuthor>
  <o:Revision>1</o:Revision>
  <o:TotalTime>2</o:TotalTime>
  <o:Created>2024-05-09T03:02:00Z</o:Created>
  <o:LastSaved>2024-05-09T03:04:00Z</o:LastSaved>
  <o:Pages>1</o:Pages>
  <o:Words>289</o:Words>
  <o:Characters>1651</o:Characters>
  <o:Lines>13</o:Lines>
  <o:Paragraphs>3</o:Paragraphs>
  <o:CharactersWithSpaces>1937</o:CharactersWithSpaces>
  <o:Version>16.00</o:Version>
 </o:DocumentProperties>
 <o:OfficeDocumentSettings>
  <o:AllowPNG/>
 </o:OfficeDocumentSettings>
</xml><![endif]-->
<link rel=themeData href="Dear_files/themedata.thmx">
<link rel=colorSchemeMapping href="Dear_files/colorschememapping.xml">
<!--[if gte mso 9]><xml>
 <w:WordDocument>
  <w:SpellingState>Clean</w:SpellingState>
  <w:GrammarState>Clean</w:GrammarState>
  <w:TrackMoves>false</w:TrackMoves>
  <w:TrackFormatting/>
  <w:PunctuationKerning/>
  <w:ValidateAgainstSchemas/>
  <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>
  <w:IgnoreMixedContent>false</w:IgnoreMixedContent>
  <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>
  <w:DoNotPromoteQF/>
  <w:LidThemeOther>EN-MY</w:LidThemeOther>
  <w:LidThemeAsian>ZH-CN</w:LidThemeAsian>
  <w:LidThemeComplexScript>X-NONE</w:LidThemeComplexScript>
  <w:Compatibility>
   <w:BreakWrappedTables/>
   <w:SnapToGridInCell/>
   <w:WrapTextWithPunct/>
   <w:UseAsianBreakRules/>
   <w:DontGrowAutofit/>
   <w:SplitPgBreakAndParaMark/>
   <w:EnableOpenTypeKerning/>
   <w:DontFlipMirrorIndents/>
   <w:OverrideTableStyleHps/>
   <w:UseFELayout/>
  </w:Compatibility>
  <m:mathPr>
   <m:mathFont m:val="Cambria Math"/>
   <m:brkBin m:val="before"/>
   <m:brkBinSub m:val="&#45;-"/>
   <m:smallFrac m:val="off"/>
   <m:dispDef/>
   <m:lMargin m:val="0"/>
   <m:rMargin m:val="0"/>
   <m:defJc m:val="centerGroup"/>
   <m:wrapIndent m:val="1440"/>
   <m:intLim m:val="subSup"/>
   <m:naryLim m:val="undOvr"/>
  </m:mathPr></w:WordDocument>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <w:LatentStyles DefLockedState="false" DefUnhideWhenUsed="false"
  DefSemiHidden="false" DefQFormat="false" DefPriority="99"
  LatentStyleCount="376">
  <w:LsdException Locked="false" Priority="0" QFormat="true" Name="Normal"/>
  <w:LsdException Locked="false" Priority="9" QFormat="true" Name="heading 1"/>
  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="heading 2"/>
  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="heading 3"/>
  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="heading 4"/>
  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="heading 5"/>
  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="heading 6"/>
  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="heading 7"/>
  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="heading 8"/>
  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="heading 9"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 5"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 6"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 7"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 8"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 9"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 1"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 2"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 3"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 4"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 5"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 6"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 7"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 8"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 9"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Normal Indent"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="footnote text"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="annotation text"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="header"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="footer"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index heading"/>
  <w:LsdException Locked="false" Priority="35" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="caption"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="table of figures"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="envelope address"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="envelope return"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="footnote reference"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="annotation reference"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="line number"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="page number"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="endnote reference"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="endnote text"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="table of authorities"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="macro"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="toa heading"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Bullet"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Number"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List 5"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Bullet 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Bullet 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Bullet 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Bullet 5"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Number 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Number 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Number 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Number 5"/>
  <w:LsdException Locked="false" Priority="10" QFormat="true" Name="Title"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Closing"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Signature"/>
  <w:LsdException Locked="false" Priority="1" SemiHidden="true"
   UnhideWhenUsed="true" Name="Default Paragraph Font"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Body Text"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Body Text Indent"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Continue"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Continue 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Continue 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Continue 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Continue 5"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Message Header"/>
  <w:LsdException Locked="false" Priority="11" QFormat="true" Name="Subtitle"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Salutation"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Date"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Body Text First Indent"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Body Text First Indent 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Note Heading"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Body Text 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Body Text 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Body Text Indent 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Body Text Indent 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Block Text"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Hyperlink"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="FollowedHyperlink"/>
  <w:LsdException Locked="false" Priority="22" QFormat="true" Name="Strong"/>
  <w:LsdException Locked="false" Priority="20" QFormat="true" Name="Emphasis"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Document Map"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Plain Text"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="E-mail Signature"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Top of Form"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Bottom of Form"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Normal (Web)"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Acronym"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Address"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Cite"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Code"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Definition"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Keyboard"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Preformatted"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Sample"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Typewriter"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Variable"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Normal Table"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="annotation subject"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="No List"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Outline List 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Outline List 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Outline List 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Simple 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Simple 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Simple 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Classic 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Classic 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Classic 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Classic 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Colorful 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Colorful 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Colorful 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Columns 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Columns 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Columns 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Columns 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Columns 5"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Grid 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Grid 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Grid 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Grid 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Grid 5"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Grid 6"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Grid 7"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Grid 8"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table List 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table List 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table List 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table List 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table List 5"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table List 6"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table List 7"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table List 8"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table 3D effects 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table 3D effects 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table 3D effects 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Contemporary"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Elegant"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Professional"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Subtle 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Subtle 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Web 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Web 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Web 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Balloon Text"/>
  <w:LsdException Locked="false" Priority="39" Name="Table Grid"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Theme"/>
  <w:LsdException Locked="false" SemiHidden="true" Name="Placeholder Text"/>
  <w:LsdException Locked="false" Priority="1" QFormat="true" Name="No Spacing"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 1"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 1"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 1"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 1"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 1"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 1"/>
  <w:LsdException Locked="false" SemiHidden="true" Name="Revision"/>
  <w:LsdException Locked="false" Priority="34" QFormat="true"
   Name="List Paragraph"/>
  <w:LsdException Locked="false" Priority="29" QFormat="true" Name="Quote"/>
  <w:LsdException Locked="false" Priority="30" QFormat="true"
   Name="Intense Quote"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 1"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 1"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 1"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 1"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 1"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 1"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 1"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 1"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 2"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 2"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 2"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 2"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 2"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 2"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 2"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 2"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 2"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 2"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 2"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 2"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 2"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 2"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 3"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 3"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 3"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 3"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 3"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 3"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 3"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 3"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 3"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 3"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 3"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 3"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 3"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 3"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 4"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 4"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 4"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 4"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 4"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 4"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 4"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 4"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 4"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 4"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 4"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 4"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 4"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 4"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 5"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 5"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 5"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 5"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 5"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 5"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 5"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 5"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 5"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 5"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 5"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 5"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 5"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 5"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 6"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 6"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 6"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 6"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 6"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 6"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 6"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 6"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 6"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 6"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 6"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 6"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 6"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 6"/>
  <w:LsdException Locked="false" Priority="19" QFormat="true"
   Name="Subtle Emphasis"/>
  <w:LsdException Locked="false" Priority="21" QFormat="true"
   Name="Intense Emphasis"/>
  <w:LsdException Locked="false" Priority="31" QFormat="true"
   Name="Subtle Reference"/>
  <w:LsdException Locked="false" Priority="32" QFormat="true"
   Name="Intense Reference"/>
  <w:LsdException Locked="false" Priority="33" QFormat="true" Name="Book Title"/>
  <w:LsdException Locked="false" Priority="37" SemiHidden="true"
   UnhideWhenUsed="true" Name="Bibliography"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="TOC Heading"/>
  <w:LsdException Locked="false" Priority="41" Name="Plain Table 1"/>
  <w:LsdException Locked="false" Priority="42" Name="Plain Table 2"/>
  <w:LsdException Locked="false" Priority="43" Name="Plain Table 3"/>
  <w:LsdException Locked="false" Priority="44" Name="Plain Table 4"/>
  <w:LsdException Locked="false" Priority="45" Name="Plain Table 5"/>
  <w:LsdException Locked="false" Priority="40" Name="Grid Table Light"/>
  <w:LsdException Locked="false" Priority="46" Name="Grid Table 1 Light"/>
  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2"/>
  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3"/>
  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4"/>
  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark"/>
  <w:LsdException Locked="false" Priority="51" Name="Grid Table 6 Colorful"/>
  <w:LsdException Locked="false" Priority="52" Name="Grid Table 7 Colorful"/>
  <w:LsdException Locked="false" Priority="46"
   Name="Grid Table 1 Light Accent 1"/>
  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 1"/>
  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 1"/>
  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 1"/>
  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 1"/>
  <w:LsdException Locked="false" Priority="51"
   Name="Grid Table 6 Colorful Accent 1"/>
  <w:LsdException Locked="false" Priority="52"
   Name="Grid Table 7 Colorful Accent 1"/>
  <w:LsdException Locked="false" Priority="46"
   Name="Grid Table 1 Light Accent 2"/>
  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 2"/>
  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 2"/>
  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 2"/>
  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 2"/>
  <w:LsdException Locked="false" Priority="51"
   Name="Grid Table 6 Colorful Accent 2"/>
  <w:LsdException Locked="false" Priority="52"
   Name="Grid Table 7 Colorful Accent 2"/>
  <w:LsdException Locked="false" Priority="46"
   Name="Grid Table 1 Light Accent 3"/>
  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 3"/>
  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 3"/>
  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 3"/>
  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 3"/>
  <w:LsdException Locked="false" Priority="51"
   Name="Grid Table 6 Colorful Accent 3"/>
  <w:LsdException Locked="false" Priority="52"
   Name="Grid Table 7 Colorful Accent 3"/>
  <w:LsdException Locked="false" Priority="46"
   Name="Grid Table 1 Light Accent 4"/>
  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 4"/>
  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 4"/>
  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 4"/>
  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 4"/>
  <w:LsdException Locked="false" Priority="51"
   Name="Grid Table 6 Colorful Accent 4"/>
  <w:LsdException Locked="false" Priority="52"
   Name="Grid Table 7 Colorful Accent 4"/>
  <w:LsdException Locked="false" Priority="46"
   Name="Grid Table 1 Light Accent 5"/>
  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 5"/>
  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 5"/>
  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 5"/>
  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 5"/>
  <w:LsdException Locked="false" Priority="51"
   Name="Grid Table 6 Colorful Accent 5"/>
  <w:LsdException Locked="false" Priority="52"
   Name="Grid Table 7 Colorful Accent 5"/>
  <w:LsdException Locked="false" Priority="46"
   Name="Grid Table 1 Light Accent 6"/>
  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 6"/>
  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 6"/>
  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 6"/>
  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 6"/>
  <w:LsdException Locked="false" Priority="51"
   Name="Grid Table 6 Colorful Accent 6"/>
  <w:LsdException Locked="false" Priority="52"
   Name="Grid Table 7 Colorful Accent 6"/>
  <w:LsdException Locked="false" Priority="46" Name="List Table 1 Light"/>
  <w:LsdException Locked="false" Priority="47" Name="List Table 2"/>
  <w:LsdException Locked="false" Priority="48" Name="List Table 3"/>
  <w:LsdException Locked="false" Priority="49" Name="List Table 4"/>
  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark"/>
  <w:LsdException Locked="false" Priority="51" Name="List Table 6 Colorful"/>
  <w:LsdException Locked="false" Priority="52" Name="List Table 7 Colorful"/>
  <w:LsdException Locked="false" Priority="46"
   Name="List Table 1 Light Accent 1"/>
  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 1"/>
  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 1"/>
  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 1"/>
  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 1"/>
  <w:LsdException Locked="false" Priority="51"
   Name="List Table 6 Colorful Accent 1"/>
  <w:LsdException Locked="false" Priority="52"
   Name="List Table 7 Colorful Accent 1"/>
  <w:LsdException Locked="false" Priority="46"
   Name="List Table 1 Light Accent 2"/>
  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 2"/>
  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 2"/>
  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 2"/>
  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 2"/>
  <w:LsdException Locked="false" Priority="51"
   Name="List Table 6 Colorful Accent 2"/>
  <w:LsdException Locked="false" Priority="52"
   Name="List Table 7 Colorful Accent 2"/>
  <w:LsdException Locked="false" Priority="46"
   Name="List Table 1 Light Accent 3"/>
  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 3"/>
  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 3"/>
  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 3"/>
  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 3"/>
  <w:LsdException Locked="false" Priority="51"
   Name="List Table 6 Colorful Accent 3"/>
  <w:LsdException Locked="false" Priority="52"
   Name="List Table 7 Colorful Accent 3"/>
  <w:LsdException Locked="false" Priority="46"
   Name="List Table 1 Light Accent 4"/>
  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 4"/>
  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 4"/>
  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 4"/>
  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 4"/>
  <w:LsdException Locked="false" Priority="51"
   Name="List Table 6 Colorful Accent 4"/>
  <w:LsdException Locked="false" Priority="52"
   Name="List Table 7 Colorful Accent 4"/>
  <w:LsdException Locked="false" Priority="46"
   Name="List Table 1 Light Accent 5"/>
  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 5"/>
  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 5"/>
  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 5"/>
  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 5"/>
  <w:LsdException Locked="false" Priority="51"
   Name="List Table 6 Colorful Accent 5"/>
  <w:LsdException Locked="false" Priority="52"
   Name="List Table 7 Colorful Accent 5"/>
  <w:LsdException Locked="false" Priority="46"
   Name="List Table 1 Light Accent 6"/>
  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 6"/>
  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 6"/>
  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 6"/>
  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 6"/>
  <w:LsdException Locked="false" Priority="51"
   Name="List Table 6 Colorful Accent 6"/>
  <w:LsdException Locked="false" Priority="52"
   Name="List Table 7 Colorful Accent 6"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Mention"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Smart Hyperlink"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Hashtag"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Unresolved Mention"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Smart Link"/>
 </w:LatentStyles>
</xml><![endif]-->
<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;
	mso-font-charset:0;
	mso-generic-font-family:roman;
	mso-font-pitch:variable;
	mso-font-signature:-536869121 1107305727 33554432 0 415 0;}
@font-face
	{font-family:DengXian;
	panose-1:2 1 6 0 3 1 1 1 1 1;
	mso-font-alt:\7B49\7EBF;
	mso-font-charset:134;
	mso-generic-font-family:auto;
	mso-font-pitch:variable;
	mso-font-signature:-1610612033 953122042 22 0 262159 0;}
@font-face
	{font-family:Aptos;
	mso-font-charset:0;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:536871559 3 0 0 415 0;}
@font-face
	{font-family:"\@DengXian";
	panose-1:2 1 6 0 3 1 1 1 1 1;
	mso-font-charset:134;
	mso-generic-font-family:auto;
	mso-font-pitch:variable;
	mso-font-signature:-1610612033 953122042 22 0 262159 0;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-parent:"";
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:8.0pt;
	margin-left:0cm;
	line-height:115%;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:DengXian;
	mso-fareast-theme-font:minor-fareast;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:minor-bidi;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;}
h1
	{mso-style-priority:9;
	mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-link:"Heading 1 Char";
	mso-style-next:Normal;
	margin-top:18.0pt;
	margin-right:0cm;
	margin-bottom:4.0pt;
	margin-left:0cm;
	line-height:115%;
	mso-pagination:widow-orphan lines-together;
	page-break-after:avoid;
	mso-outline-level:1;
	font-size:20.0pt;
	font-family:"Aptos Display",sans-serif;
	mso-ascii-font-family:"Aptos Display";
	mso-ascii-theme-font:major-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:"Aptos Display";
	mso-hansi-theme-font:major-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#0F4761;
	mso-themecolor:accent1;
	mso-themeshade:191;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;
	font-weight:normal;}
h2
	{mso-style-noshow:yes;
	mso-style-priority:9;
	mso-style-qformat:yes;
	mso-style-link:"Heading 2 Char";
	mso-style-next:Normal;
	margin-top:8.0pt;
	margin-right:0cm;
	margin-bottom:4.0pt;
	margin-left:0cm;
	line-height:115%;
	mso-pagination:widow-orphan lines-together;
	page-break-after:avoid;
	mso-outline-level:2;
	font-size:16.0pt;
	font-family:"Aptos Display",sans-serif;
	mso-ascii-font-family:"Aptos Display";
	mso-ascii-theme-font:major-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:"Aptos Display";
	mso-hansi-theme-font:major-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#0F4761;
	mso-themecolor:accent1;
	mso-themeshade:191;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;
	font-weight:normal;}
h3
	{mso-style-noshow:yes;
	mso-style-priority:9;
	mso-style-qformat:yes;
	mso-style-link:"Heading 3 Char";
	mso-style-next:Normal;
	margin-top:8.0pt;
	margin-right:0cm;
	margin-bottom:4.0pt;
	margin-left:0cm;
	line-height:115%;
	mso-pagination:widow-orphan lines-together;
	page-break-after:avoid;
	mso-outline-level:3;
	font-size:14.0pt;
	font-family:"Aptos",sans-serif;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#0F4761;
	mso-themecolor:accent1;
	mso-themeshade:191;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;
	font-weight:normal;}
h4
	{mso-style-noshow:yes;
	mso-style-priority:9;
	mso-style-qformat:yes;
	mso-style-link:"Heading 4 Char";
	mso-style-next:Normal;
	margin-top:4.0pt;
	margin-right:0cm;
	margin-bottom:2.0pt;
	margin-left:0cm;
	line-height:115%;
	mso-pagination:widow-orphan lines-together;
	page-break-after:avoid;
	mso-outline-level:4;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#0F4761;
	mso-themecolor:accent1;
	mso-themeshade:191;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;
	font-weight:normal;
	font-style:italic;}
h5
	{mso-style-noshow:yes;
	mso-style-priority:9;
	mso-style-qformat:yes;
	mso-style-link:"Heading 5 Char";
	mso-style-next:Normal;
	margin-top:4.0pt;
	margin-right:0cm;
	margin-bottom:2.0pt;
	margin-left:0cm;
	line-height:115%;
	mso-pagination:widow-orphan lines-together;
	page-break-after:avoid;
	mso-outline-level:5;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#0F4761;
	mso-themecolor:accent1;
	mso-themeshade:191;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;
	font-weight:normal;}
h6
	{mso-style-noshow:yes;
	mso-style-priority:9;
	mso-style-qformat:yes;
	mso-style-link:"Heading 6 Char";
	mso-style-next:Normal;
	margin-top:2.0pt;
	margin-right:0cm;
	margin-bottom:0cm;
	margin-left:0cm;
	line-height:115%;
	mso-pagination:widow-orphan lines-together;
	page-break-after:avoid;
	mso-outline-level:6;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#595959;
	mso-themecolor:text1;
	mso-themetint:166;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;
	font-weight:normal;
	font-style:italic;}
p.MsoHeading7, li.MsoHeading7, div.MsoHeading7
	{mso-style-noshow:yes;
	mso-style-priority:9;
	mso-style-qformat:yes;
	mso-style-link:"Heading 7 Char";
	mso-style-next:Normal;
	margin-top:2.0pt;
	margin-right:0cm;
	margin-bottom:0cm;
	margin-left:0cm;
	line-height:115%;
	mso-pagination:widow-orphan lines-together;
	page-break-after:avoid;
	mso-outline-level:7;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#595959;
	mso-themecolor:text1;
	mso-themetint:166;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;}
p.MsoHeading8, li.MsoHeading8, div.MsoHeading8
	{mso-style-noshow:yes;
	mso-style-priority:9;
	mso-style-qformat:yes;
	mso-style-link:"Heading 8 Char";
	mso-style-next:Normal;
	margin:0cm;
	line-height:115%;
	mso-pagination:widow-orphan lines-together;
	page-break-after:avoid;
	mso-outline-level:8;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#272727;
	mso-themecolor:text1;
	mso-themetint:216;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;
	font-style:italic;}
p.MsoHeading9, li.MsoHeading9, div.MsoHeading9
	{mso-style-noshow:yes;
	mso-style-priority:9;
	mso-style-qformat:yes;
	mso-style-link:"Heading 9 Char";
	mso-style-next:Normal;
	margin:0cm;
	line-height:115%;
	mso-pagination:widow-orphan lines-together;
	page-break-after:avoid;
	mso-outline-level:9;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#272727;
	mso-themecolor:text1;
	mso-themetint:216;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;}
p.MsoTitle, li.MsoTitle, div.MsoTitle
	{mso-style-priority:10;
	mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-link:"Title Char";
	mso-style-next:Normal;
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:4.0pt;
	margin-left:0cm;
	mso-add-space:auto;
	mso-pagination:widow-orphan;
	font-size:28.0pt;
	font-family:"Aptos Display",sans-serif;
	mso-ascii-font-family:"Aptos Display";
	mso-ascii-theme-font:major-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:"Aptos Display";
	mso-hansi-theme-font:major-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	letter-spacing:-.5pt;
	mso-font-kerning:14.0pt;
	mso-ligatures:standardcontextual;}
p.MsoTitleCxSpFirst, li.MsoTitleCxSpFirst, div.MsoTitleCxSpFirst
	{mso-style-priority:10;
	mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-link:"Title Char";
	mso-style-next:Normal;
	mso-style-type:export-only;
	margin:0cm;
	mso-add-space:auto;
	mso-pagination:widow-orphan;
	font-size:28.0pt;
	font-family:"Aptos Display",sans-serif;
	mso-ascii-font-family:"Aptos Display";
	mso-ascii-theme-font:major-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:"Aptos Display";
	mso-hansi-theme-font:major-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	letter-spacing:-.5pt;
	mso-font-kerning:14.0pt;
	mso-ligatures:standardcontextual;}
p.MsoTitleCxSpMiddle, li.MsoTitleCxSpMiddle, div.MsoTitleCxSpMiddle
	{mso-style-priority:10;
	mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-link:"Title Char";
	mso-style-next:Normal;
	mso-style-type:export-only;
	margin:0cm;
	mso-add-space:auto;
	mso-pagination:widow-orphan;
	font-size:28.0pt;
	font-family:"Aptos Display",sans-serif;
	mso-ascii-font-family:"Aptos Display";
	mso-ascii-theme-font:major-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:"Aptos Display";
	mso-hansi-theme-font:major-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	letter-spacing:-.5pt;
	mso-font-kerning:14.0pt;
	mso-ligatures:standardcontextual;}
p.MsoTitleCxSpLast, li.MsoTitleCxSpLast, div.MsoTitleCxSpLast
	{mso-style-priority:10;
	mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-link:"Title Char";
	mso-style-next:Normal;
	mso-style-type:export-only;
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:4.0pt;
	margin-left:0cm;
	mso-add-space:auto;
	mso-pagination:widow-orphan;
	font-size:28.0pt;
	font-family:"Aptos Display",sans-serif;
	mso-ascii-font-family:"Aptos Display";
	mso-ascii-theme-font:major-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:"Aptos Display";
	mso-hansi-theme-font:major-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	letter-spacing:-.5pt;
	mso-font-kerning:14.0pt;
	mso-ligatures:standardcontextual;}
p.MsoSubtitle, li.MsoSubtitle, div.MsoSubtitle
	{mso-style-priority:11;
	mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-link:"Subtitle Char";
	mso-style-next:Normal;
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:8.0pt;
	margin-left:0cm;
	line-height:115%;
	mso-pagination:widow-orphan;
	font-size:14.0pt;
	font-family:"Aptos",sans-serif;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#595959;
	mso-themecolor:text1;
	mso-themetint:166;
	letter-spacing:.75pt;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;}
a:link, span.MsoHyperlink
	{mso-style-noshow:yes;
	mso-style-priority:99;
	color:blue;
	text-decoration:underline;
	text-underline:single;}
a:visited, span.MsoHyperlinkFollowed
	{mso-style-noshow:yes;
	mso-style-priority:99;
	color:#96607D;
	mso-themecolor:followedhyperlink;
	text-decoration:underline;
	text-underline:single;}
p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph
	{mso-style-priority:34;
	mso-style-unhide:no;
	mso-style-qformat:yes;
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:8.0pt;
	margin-left:36.0pt;
	mso-add-space:auto;
	line-height:115%;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:DengXian;
	mso-fareast-theme-font:minor-fareast;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:minor-bidi;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;}
p.MsoListParagraphCxSpFirst, li.MsoListParagraphCxSpFirst, div.MsoListParagraphCxSpFirst
	{mso-style-priority:34;
	mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-type:export-only;
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:0cm;
	margin-left:36.0pt;
	mso-add-space:auto;
	line-height:115%;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:DengXian;
	mso-fareast-theme-font:minor-fareast;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:minor-bidi;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;}
p.MsoListParagraphCxSpMiddle, li.MsoListParagraphCxSpMiddle, div.MsoListParagraphCxSpMiddle
	{mso-style-priority:34;
	mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-type:export-only;
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:0cm;
	margin-left:36.0pt;
	mso-add-space:auto;
	line-height:115%;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:DengXian;
	mso-fareast-theme-font:minor-fareast;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:minor-bidi;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;}
p.MsoListParagraphCxSpLast, li.MsoListParagraphCxSpLast, div.MsoListParagraphCxSpLast
	{mso-style-priority:34;
	mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-type:export-only;
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:8.0pt;
	margin-left:36.0pt;
	mso-add-space:auto;
	line-height:115%;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:DengXian;
	mso-fareast-theme-font:minor-fareast;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:minor-bidi;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;}
p.MsoQuote, li.MsoQuote, div.MsoQuote
	{mso-style-priority:29;
	mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-link:"Quote Char";
	mso-style-next:Normal;
	margin-top:8.0pt;
	margin-right:0cm;
	margin-bottom:8.0pt;
	margin-left:0cm;
	text-align:center;
	line-height:115%;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:DengXian;
	mso-fareast-theme-font:minor-fareast;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:minor-bidi;
	color:#404040;
	mso-themecolor:text1;
	mso-themetint:191;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;
	font-style:italic;}
p.MsoIntenseQuote, li.MsoIntenseQuote, div.MsoIntenseQuote
	{mso-style-priority:30;
	mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-link:"Intense Quote Char";
	mso-style-next:Normal;
	margin-top:18.0pt;
	margin-right:43.2pt;
	margin-bottom:18.0pt;
	margin-left:43.2pt;
	text-align:center;
	line-height:115%;
	mso-pagination:widow-orphan;
	border:none;
	mso-border-top-alt:solid #0F4761 .5pt;
	mso-border-top-themecolor:accent1;
	mso-border-top-themeshade:191;
	mso-border-bottom-alt:solid #0F4761 .5pt;
	mso-border-bottom-themecolor:accent1;
	mso-border-bottom-themeshade:191;
	padding:0cm;
	mso-padding-alt:10.0pt 0cm 10.0pt 0cm;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:DengXian;
	mso-fareast-theme-font:minor-fareast;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:minor-bidi;
	color:#0F4761;
	mso-themecolor:accent1;
	mso-themeshade:191;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;
	font-style:italic;}
span.MsoIntenseEmphasis
	{mso-style-priority:21;
	mso-style-unhide:no;
	mso-style-qformat:yes;
	color:#0F4761;
	mso-themecolor:accent1;
	mso-themeshade:191;
	font-style:italic;}
span.MsoIntenseReference
	{mso-style-priority:32;
	mso-style-unhide:no;
	mso-style-qformat:yes;
	font-variant:small-caps;
	color:#0F4761;
	mso-themecolor:accent1;
	mso-themeshade:191;
	letter-spacing:.25pt;
	font-weight:bold;}
span.Heading1Char
	{mso-style-name:"Heading 1 Char";
	mso-style-priority:9;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:"Heading 1";
	mso-ansi-font-size:20.0pt;
	mso-bidi-font-size:20.0pt;
	font-family:"Aptos Display",sans-serif;
	mso-ascii-font-family:"Aptos Display";
	mso-ascii-theme-font:major-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:"Aptos Display";
	mso-hansi-theme-font:major-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#0F4761;
	mso-themecolor:accent1;
	mso-themeshade:191;}
span.Heading2Char
	{mso-style-name:"Heading 2 Char";
	mso-style-noshow:yes;
	mso-style-priority:9;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:"Heading 2";
	mso-ansi-font-size:16.0pt;
	mso-bidi-font-size:16.0pt;
	font-family:"Aptos Display",sans-serif;
	mso-ascii-font-family:"Aptos Display";
	mso-ascii-theme-font:major-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:"Aptos Display";
	mso-hansi-theme-font:major-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#0F4761;
	mso-themecolor:accent1;
	mso-themeshade:191;}
span.Heading3Char
	{mso-style-name:"Heading 3 Char";
	mso-style-noshow:yes;
	mso-style-priority:9;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:"Heading 3";
	mso-ansi-font-size:14.0pt;
	mso-bidi-font-size:14.0pt;
	font-family:"DengXian Light";
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#0F4761;
	mso-themecolor:accent1;
	mso-themeshade:191;}
span.Heading4Char
	{mso-style-name:"Heading 4 Char";
	mso-style-noshow:yes;
	mso-style-priority:9;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:"Heading 4";
	font-family:"DengXian Light";
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#0F4761;
	mso-themecolor:accent1;
	mso-themeshade:191;
	font-style:italic;}
span.Heading5Char
	{mso-style-name:"Heading 5 Char";
	mso-style-noshow:yes;
	mso-style-priority:9;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:"Heading 5";
	font-family:"DengXian Light";
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#0F4761;
	mso-themecolor:accent1;
	mso-themeshade:191;}
span.Heading6Char
	{mso-style-name:"Heading 6 Char";
	mso-style-noshow:yes;
	mso-style-priority:9;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:"Heading 6";
	font-family:"DengXian Light";
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#595959;
	mso-themecolor:text1;
	mso-themetint:166;
	font-style:italic;}
span.Heading7Char
	{mso-style-name:"Heading 7 Char";
	mso-style-noshow:yes;
	mso-style-priority:9;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:"Heading 7";
	font-family:"DengXian Light";
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#595959;
	mso-themecolor:text1;
	mso-themetint:166;}
span.Heading8Char
	{mso-style-name:"Heading 8 Char";
	mso-style-noshow:yes;
	mso-style-priority:9;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:"Heading 8";
	font-family:"DengXian Light";
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#272727;
	mso-themecolor:text1;
	mso-themetint:216;
	font-style:italic;}
span.Heading9Char
	{mso-style-name:"Heading 9 Char";
	mso-style-noshow:yes;
	mso-style-priority:9;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:"Heading 9";
	font-family:"DengXian Light";
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#272727;
	mso-themecolor:text1;
	mso-themetint:216;}
span.TitleChar
	{mso-style-name:"Title Char";
	mso-style-priority:10;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:Title;
	mso-ansi-font-size:28.0pt;
	mso-bidi-font-size:28.0pt;
	font-family:"Aptos Display",sans-serif;
	mso-ascii-font-family:"Aptos Display";
	mso-ascii-theme-font:major-latin;
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-hansi-font-family:"Aptos Display";
	mso-hansi-theme-font:major-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	letter-spacing:-.5pt;
	mso-font-kerning:14.0pt;}
span.SubtitleChar
	{mso-style-name:"Subtitle Char";
	mso-style-priority:11;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:Subtitle;
	mso-ansi-font-size:14.0pt;
	mso-bidi-font-size:14.0pt;
	font-family:"DengXian Light";
	mso-fareast-font-family:"DengXian Light";
	mso-fareast-theme-font:major-fareast;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:major-bidi;
	color:#595959;
	mso-themecolor:text1;
	mso-themetint:166;
	letter-spacing:.75pt;}
span.QuoteChar
	{mso-style-name:"Quote Char";
	mso-style-priority:29;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:Quote;
	color:#404040;
	mso-themecolor:text1;
	mso-themetint:191;
	font-style:italic;}
span.IntenseQuoteChar
	{mso-style-name:"Intense Quote Char";
	mso-style-priority:30;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:"Intense Quote";
	color:#0F4761;
	mso-themecolor:accent1;
	mso-themeshade:191;
	font-style:italic;}
span.gmailsignatureprefix
	{mso-style-name:gmail_signature_prefix;
	mso-style-unhide:no;}
span.SpellE
	{mso-style-name:"";
	mso-spl-e:yes;}
span.GramE
	{mso-style-name:"";
	mso-gram-e:yes;}
.MsoChpDefault
	{mso-style-type:export-only;
	mso-default-props:yes;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:DengXian;
	mso-fareast-theme-font:minor-fareast;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:minor-bidi;}
.MsoPapDefault
	{mso-style-type:export-only;
	margin-bottom:8.0pt;
	line-height:115%;}
@page WordSection1
	{size:595.3pt 841.9pt;
	margin:72.0pt 72.0pt 72.0pt 72.0pt;
	mso-header-margin:35.4pt;
	mso-footer-margin:35.4pt;
	mso-paper-source:0;}
div.WordSection1
	{page:WordSection1;}
-->
</style>
<!--[if gte mso 10]>
<style>
 /* Style Definitions */
 table.MsoNormalTable
	{mso-style-name:"Table Normal";
	mso-tstyle-rowband-size:0;
	mso-tstyle-colband-size:0;
	mso-style-noshow:yes;
	mso-style-priority:99;
	mso-style-parent:"";
	mso-padding-alt:0cm 5.4pt 0cm 5.4pt;
	mso-para-margin-top:0cm;
	mso-para-margin-right:0cm;
	mso-para-margin-bottom:8.0pt;
	mso-para-margin-left:0cm;
	line-height:115%;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;
	mso-ascii-font-family:Aptos;
	mso-ascii-theme-font:minor-latin;
	mso-hansi-font-family:Aptos;
	mso-hansi-theme-font:minor-latin;
	mso-font-kerning:1.0pt;
	mso-ligatures:standardcontextual;}
</style>
<![endif]--><!--[if gte mso 9]><xml>
 <o:shapedefaults v:ext="edit" spidmax="1026"/>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout></xml><![endif]-->
</head>

<body lang=EN-MY link=blue vlink="#96607D" style='tab-interval:36.0pt;
word-wrap:break-word'>

<div class=WordSection1>

<p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
mso-font-kerning:0pt;mso-ligatures:none'>Dear """

        altbody2="""
        ,<o:p></o:p></span></p>
        
        <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
mso-font-kerning:0pt;mso-ligatures:none'><o:p>&nbsp;</o:p></span></p>

<p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
mso-font-kerning:0pt;mso-ligatures:none'>Greetings from Agents of Tech!<o:p></o:p></span></p>

<p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
mso-font-kerning:0pt;mso-ligatures:none'><o:p>&nbsp;</o:p></span></p>

<p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
mso-font-kerning:0pt;mso-ligatures:none'>Thank you for joining us at <span
class=SpellE>NexTech</span> Conference &amp; Expo 2024! We hope that the
knowledge <span class=GramE>gained</span> and connections made during <span
class=SpellE>NexTech</span> Conference &amp; Expo 2024 will continue to inspire
and empower you in your endeavours.<o:p></o:p></span></p>

<p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
mso-font-kerning:0pt;mso-ligatures:none'><o:p>&nbsp;</o:p></span></p>

<p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
mso-font-kerning:0pt;mso-ligatures:none'>You may find your e-certificate
awarded for your participation attached to this email. Please let us know
immediately if there is any error with the certificates. <o:p></o:p></span></p>

<p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
mso-font-kerning:0pt;mso-ligatures:none'><o:p>&nbsp;</o:p></span></p>

<p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
mso-font-kerning:0pt;mso-ligatures:none'>Once again, thank you for your
support. We look forward to welcoming you back in our future events!<o:p></o:p></span></p>

<p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
mso-font-kerning:0pt;mso-ligatures:none'><o:p>&nbsp;</o:p></span></p>

<p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
mso-font-kerning:0pt;mso-ligatures:none'>Follow us on our IG @agentsoftech.tlc
for future exciting events and updates!<o:p></o:p></span></p>

<p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
mso-font-kerning:0pt;mso-ligatures:none'>-- <o:p></o:p></span></p>

<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=480
 style='width:360.0pt;border-collapse:collapse;mso-yfti-tbllook:1184;
 mso-padding-alt:0cm 0cm 0cm 0cm'>
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
  <td width=160 valign=top style='width:120.0pt;padding:7.5pt 0cm 9.0pt 0cm'>
  <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><i><span
  style='font-size:10.0pt;font-family:"Arial",sans-serif;mso-fareast-font-family:
  "Times New Roman";color:#444444;mso-font-kerning:0pt;mso-ligatures:none'>Best
  Regards,</span></i><span style='font-size:10.0pt;font-family:"Arial",sans-serif;
  mso-fareast-font-family:"Times New Roman";color:#444444;mso-font-kerning:
  0pt;mso-ligatures:none'><br>
  <span style='mso-no-proof:yes'><!--[if gte vml 1]><v:shapetype id="_x0000_t75"
   coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe"
   filled="f" stroked="f">
   <v:stroke joinstyle="miter"/>
   <v:formulas>
    <v:f eqn="if lineDrawn pixelLineWidth 0"/>
    <v:f eqn="sum @0 1 0"/>
    <v:f eqn="sum 0 0 @1"/>
    <v:f eqn="prod @2 1 2"/>
    <v:f eqn="prod @3 21600 pixelWidth"/>
    <v:f eqn="prod @3 21600 pixelHeight"/>
    <v:f eqn="sum @0 0 1"/>
    <v:f eqn="prod @6 1 2"/>
    <v:f eqn="prod @7 21600 pixelWidth"/>
    <v:f eqn="sum @8 21600 0"/>
    <v:f eqn="prod @7 21600 pixelHeight"/>
    <v:f eqn="sum @10 21600 0"/>
   </v:formulas>
   <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
   <o:lock v:ext="edit" aspectratio="t"/>
  </v:shapetype><v:shape id="Picture_x0020_4" o:spid="_x0000_i1028" type="#_x0000_t75"
   alt="A blue lines with circles and lines on a black background&#10;&#10;Description automatically generated"
   style='width:150pt;height:150pt;visibility:visible;mso-wrap-style:square'>
   <v:imagedata src="https://ci3.googleusercontent.com/mail-sig/AIorK4x1K3sHIxLHv_RWMCANHDo3qfDvIrlYts2xguT9xxtIHB3WV8V6G0TpRFAiGREGXNuym2Y5IEc " o:title="A blue lines with circles and lines on a black background&#10;&#10;Description automatically generated"/>
  </v:shape><![endif]--><![if !vml]><img width=200 height=200
  src="https://ci3.googleusercontent.com/mail-sig/AIorK4x1K3sHIxLHv_RWMCANHDo3qfDvIrlYts2xguT9xxtIHB3WV8V6G0TpRFAiGREGXNuym2Y5IEc "
  alt="A blue lines with circles and lines on a black background&#10;&#10;Description automatically generated"
  v:shapes="Picture_x0020_4"><![endif]></span><o:p></o:p></span></p>
  </td>
  <td width=320 style='width:240.0pt;padding:4.5pt 0cm 4.5pt 0cm'>
  <table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0
   style='border-collapse:collapse;mso-yfti-tbllook:1184;mso-padding-alt:0cm 0cm 0cm 0cm'>
   <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
    <td style='padding:0cm 0cm 0cm 0cm'>
    <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><b><span
    style='font-size:13.5pt;font-family:"Times New Roman",serif;mso-fareast-font-family:
    "Times New Roman";color:#3D3C3F;mso-font-kerning:0pt;mso-ligatures:none'><br>
    <br>
    Taylor's Agents of Tech Club</span></b><b><span style='font-family:"Times New Roman",serif;
    mso-fareast-font-family:"Times New Roman";color:#3D3C3F;mso-font-kerning:
    0pt;mso-ligatures:none'><o:p></o:p></span></b></p>
    </td>
   </tr>
   <tr style='mso-yfti-irow:1'>
    <td style='padding:0cm 0cm 8.25pt 0cm'></td>
   </tr>
   <tr style='mso-yfti-irow:2'>
    <td style='padding:0cm 0cm 0cm 0cm'>
    <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
    style='font-size:10.0pt;font-family:"Times New Roman",serif;mso-fareast-font-family:
    "Times New Roman";color:black;mso-font-kerning:0pt;mso-ligatures:none'>Mobile:&nbsp;<a
    href="tel:++6016-4598506" target="_blank"><span style='color:#1155CC'>+6010-5633381</span></a></span><span
    style='font-size:10.0pt;font-family:"Times New Roman",serif;mso-fareast-font-family:
    "Times New Roman";mso-font-kerning:0pt;mso-ligatures:none'><o:p></o:p></span></p>
    </td>
   </tr>
   <tr style='mso-yfti-irow:3'>
    <td style='padding:0cm 0cm 0cm 0cm'>
    <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
    style='font-size:10.0pt;font-family:"Times New Roman",serif;mso-fareast-font-family:
    "Times New Roman";color:black;mso-font-kerning:0pt;mso-ligatures:none'>Email:&nbsp;<a
    href="mailto:info@agentsoftech.my" target="_blank"><span style='color:#1155CC'>info@agentsoftech.my</span></a></span><span
    style='font-size:10.0pt;font-family:"Times New Roman",serif;mso-fareast-font-family:
    "Times New Roman";mso-font-kerning:0pt;mso-ligatures:none'><br>
    <span style='color:black'>Website:&nbsp;</span><a
    href="https://www.agentsoftech.my/" target="_blank"><span style='color:
    #3D85C6'>Taylor's Agents of Tech Official Website</span></a><o:p></o:p></span></p>
    </td>
   </tr>
   <tr style='mso-yfti-irow:4'>
    <td style='padding:0cm 0cm 0cm 0cm'>
    <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><b><span
    style='font-size:10.0pt;font-family:"Times New Roman",serif;mso-fareast-font-family:
    "Times New Roman";color:black;mso-font-kerning:0pt;mso-ligatures:none'>Taylor's
    University Lakeside Campus</span></b><span style='font-size:10.0pt;
    font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
    color:black;mso-font-kerning:0pt;mso-ligatures:none'><br>
    No 1, Jalan Taylor's, 47500 Subang Jaya,</span><span style='font-size:10.0pt;
    font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
    color:#9B9B9B;mso-font-kerning:0pt;mso-ligatures:none'><o:p></o:p></span></p>
    </td>
   </tr>
   <tr style='mso-yfti-irow:5'>
    <td style='padding:0cm 0cm 0cm 0cm'>
    <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><span
    style='font-size:10.0pt;font-family:"Times New Roman",serif;mso-fareast-font-family:
    "Times New Roman";color:black;mso-font-kerning:0pt;mso-ligatures:none'>Selangor
    Darul Ehsan, Malaysia</span><span style='font-size:10.0pt;font-family:"Times New Roman",serif;
    mso-fareast-font-family:"Times New Roman";color:#9B9B9B;mso-font-kerning:
    0pt;mso-ligatures:none'><o:p></o:p></span></p>
    </td>
   </tr>
   <tr style='mso-yfti-irow:6;mso-yfti-lastrow:yes'>
    <td style='padding:4.5pt 0cm 0cm 0cm'>
    <p class=MsoNormal style='margin-bottom:0cm;line-height:normal'><a
    href="https://www.facebook.com/AgentsOfTech" target="_blank"><span
    style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
    color:#337AB7;mso-font-kerning:0pt;mso-no-proof:yes;text-decoration:none;
    text-underline:none'><!--[if gte vml 1]><v:shape id="Picture_x0020_3"
     o:spid="_x0000_i1027" type="#_x0000_t75" alt="Facebook icon"
     href="https://www.facebook.com/AgentsOfTech" target="&quot;_blank&quot;"
     style='width:15pt;height:15pt;visibility:visible;mso-wrap-style:square'
     o:button="t">
     <v:imagedata src="https://ci3.googleusercontent.com/meips/ADKq_NYmDXiOYv3e35S9R_w8Q5L7MCXbeSEBPmxTyRyqT77TEXV6Y7cc651RG7H2B1wKsxM7omClHGQ3T8E_Wc8dSp1a7nYB4V2ZfV99ysqf2iL4m5pJbAvv2WiLctTHveSjDZXNHmgQ2b8CeLLCUVBY=s0-d-e1-ft#https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/elegant-logo/fb.png" o:title="Facebook icon"/>
    </v:shape><![endif]--><![if !vml]><span style='mso-ignore:vglayout'><img
    border=0 width=20 height=20 src="https://ci3.googleusercontent.com/meips/ADKq_NYmDXiOYv3e35S9R_w8Q5L7MCXbeSEBPmxTyRyqT77TEXV6Y7cc651RG7H2B1wKsxM7omClHGQ3T8E_Wc8dSp1a7nYB4V2ZfV99ysqf2iL4m5pJbAvv2WiLctTHveSjDZXNHmgQ2b8CeLLCUVBY=s0-d-e1-ft#https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/elegant-logo/fb.png"
    alt="Facebook icon" v:shapes="Picture_x0020_3"></span><![endif]></span></a><span
    style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
    mso-font-kerning:0pt;mso-ligatures:none'>&nbsp;&nbsp;</span><a
    href="https://www.linkedin.com/company/tayloraot" target="_blank"><span
    style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
    color:#337AB7;mso-font-kerning:0pt;mso-no-proof:yes;text-decoration:none;
    text-underline:none'><!--[if gte vml 1]><v:shape id="Picture_x0020_2"
     o:spid="_x0000_i1026" type="#_x0000_t75" alt="LinkedIn icon"
     href="https://www.linkedin.com/company/tayloraot" target="&quot;_blank&quot;"
     style='width:15pt;height:15pt;visibility:visible;mso-wrap-style:square'
     o:button="t">
     <v:imagedata src="https://ci3.googleusercontent.com/meips/ADKq_NYC42pOpZWR3HqN-k61FxuUbhR-wi9db2GkXWMe2hYaKX253OSKJ1NyjiNbF_dckLDXf0HhRZsr_BqoL0Yw-rdwz5mNMf8oKimVwUA0VBSgLmfZxkMXSN9PYZmDQoSYvJbaAmVzb25CA9PvhNMn=s0-d-e1-ft#https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/elegant-logo/ln.png" o:title="LinkedIn icon"/>
    </v:shape><![endif]--><![if !vml]><span style='mso-ignore:vglayout'><img
    border=0 width=20 height=20 src="https://ci3.googleusercontent.com/meips/ADKq_NYC42pOpZWR3HqN-k61FxuUbhR-wi9db2GkXWMe2hYaKX253OSKJ1NyjiNbF_dckLDXf0HhRZsr_BqoL0Yw-rdwz5mNMf8oKimVwUA0VBSgLmfZxkMXSN9PYZmDQoSYvJbaAmVzb25CA9PvhNMn=s0-d-e1-ft#https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/elegant-logo/ln.png"
    alt="LinkedIn icon" v:shapes="Picture_x0020_2"></span><![endif]></span></a><span
    style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
    mso-font-kerning:0pt;mso-ligatures:none'>&nbsp;&nbsp;</span><a
    href="https://instagram.com/agentsoftech.tlc" target="_blank"><span
    style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
    color:#337AB7;mso-font-kerning:0pt;mso-no-proof:yes;text-decoration:none;
    text-underline:none'><!--[if gte vml 1]><v:shape id="Picture_x0020_1"
     o:spid="_x0000_i1025" type="#_x0000_t75" alt="Instagram icon"
     href="https://instagram.com/agentsoftech.tlc" target="&quot;_blank&quot;"
     style='width:15pt;height:15pt;visibility:visible;mso-wrap-style:square'
     o:button="t">
     <v:imagedata src="https://ci3.googleusercontent.com/meips/ADKq_NYLUaL4zGtazgjj5-mskIwqDmpsJNPzAL2647-muLanv4MgNUIzHUkLaUkaVZx1_bhv_j3idnqfyY0QBL4ddKKsEZINwf5q6IXfWtbHFOSaRz2PXijyorrTMi_W-QHDtk66p6JDYkfU9iVEZC8K=s0-d-e1-ft#https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/elegant-logo/it.png" o:title="Instagram icon"/>
    </v:shape><![endif]--><![if !vml]><span style='mso-ignore:vglayout'><img
    border=0 width=20 height=20 src="https://ci3.googleusercontent.com/meips/ADKq_NYLUaL4zGtazgjj5-mskIwqDmpsJNPzAL2647-muLanv4MgNUIzHUkLaUkaVZx1_bhv_j3idnqfyY0QBL4ddKKsEZINwf5q6IXfWtbHFOSaRz2PXijyorrTMi_W-QHDtk66p6JDYkfU9iVEZC8K=s0-d-e1-ft#https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/elegant-logo/it.png"
    alt="Instagram icon" v:shapes="Picture_x0020_1"></span><![endif]></span></a><span
    style='font-family:"Times New Roman",serif;mso-fareast-font-family:"Times New Roman";
    mso-font-kerning:0pt;mso-ligatures:none'>&nbsp;&nbsp;<o:p></o:p></span></p>
    </td>
   </tr>
  </table>
  </td>
 </tr>
 <tr style='mso-yfti-irow:1;mso-yfti-lastrow:yes'>
  <td width=480 colspan=2 style='width:360.0pt;border:none;border-top:solid #1594D4 1.0pt;
  mso-border-top-alt:solid #1594D4 .75pt;padding:6.0pt 0cm 0cm 0cm'>
  <p class=MsoNormal style='margin-bottom:0cm;text-align:justify;text-justify:
  inter-ideograph;line-height:normal'><span style='font-size:10.0pt;font-family:
  "Arial",sans-serif;mso-fareast-font-family:"Times New Roman";color:black;
  mso-font-kerning:0pt;mso-ligatures:none'>Confidentiality Note: This message
  (including any attachments) is intended only for the use of the individual or
  entity to which it is addressed and may contain information that is
  non-public, proprietary, privileged, confidential, and exempt from disclosure
  under applicable law or may constitute as an attorney work product. If you
  are not the intended recipient, you are hereby notified that any use,
  dissemination, distribution, or copying of this communication is strictly
  prohibited. If you have received this communication in error, notify us
  immediately by telephone and (<span class=SpellE>i</span>) destroy this
  message if a facsimile or (ii) delete this message immediately if this is an
  electronic communication.</span><span style='font-size:10.0pt;font-family:
  "Arial",sans-serif;mso-fareast-font-family:"Times New Roman";color:#9B9B9B;
  mso-font-kerning:0pt;mso-ligatures:none'><o:p></o:p></span></p>
  </td>
 </tr>
</table>

<p class=MsoNormal><o:p>&nbsp;</o:p></p>

</div>

</body>

</html>
"""
        altbody = altbody + name + altbody2
        msg.add_header('Content-Type', 'text/html')
        msg.set_content(altbody, subtype = 'html')
        
        with open(pdf_path, 'rb') as fp:
            pdf_data = fp.read()
         
        msg.add_attachment(pdf_data, maintype='application',
                                     subtype='pdf', filename=os.path.basename(pdf_path))
        
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(self.gmailFrom, self.app_password)
            s.send_message(msg)
            
    def sendBulkMessage(self, csv_file, name_column, name_column2, email_column):
        with open(csv_file, 'r', newline='', encoding='utf-8') as csvfile:
            print("Starting...")
            csvreader = csv.DictReader(csvfile)
            df = pd.read_csv(csv_file)
            num = len(df)
            i = 0
            for row in csvreader:
                i += 1
                # Join first name and last name columns
                name = ' '.join([row[name_column], row[name_column2]])
                recipient_email = row[email_column]
                # Send message to each recipient
                self.sendMessage(name, recipient_email)
                print(f"Completed {i} rows out of {num} rows")
  
def start(certificate_jpg_file, csv_file, text_pos_tuple, email, app_password, text_anchor = "mb", font_color_RGB_tuple = (0,0,0), text_alignment = 'center', font_path="None", font_size = "None"):
    Automate = Message()

    Automate.font_path = "C:\\Users\\yapzh\\Downloads\\Montserrat-Regular.ttf" # Select your custom font
     
    Automate.trueType = ImageFont.truetype(Automate.font_path, 60) # Change the number to adjust font size

    Automate.RESOURCE_ROOT = str(os.path.dirname(os.path.abspath(__file__))) # Edit this for the directory of the files

    Automate.setJPGPath("NexTech.jpg") # Edit this to the name of your Certificate image

    Automate.gmailFrom = email

    Automate.app_password = app_password

    Automate.setTextSettings((875, 570), 'mb', (255, 255, 255), 'center')

    csvfile = "C:\\Users\\yapzh\\Downloads\\Cert-AOT.csv"

    #Automate.sendBulkMessage(csvfile, 'First Name', 'Last Name', 'Email')ast.literal_eval(s)

    Automate.sendMessage("Chen Foong Lim", 'yapzhehin@gmail.com')

if __name__ == "__main__":
    args = sys.argv[1::]
    certificate_jpg_file = None
    csv_file = None
    text_pos_tuple = None
    text_anchor = "mb"
    font_color_RGB_tuple = (0,0,0)
    text_alignment = 'center'
    font_path="None"
    font_size = 30
    if len(sys.argv) > 1 and sys.argv[1] in ('-h', '--help'):
        print("Usage: python aot.py [certificate_jpg_file] [csv_file] [text_pos_tuple] [font_color_RGB_tuple=value]")
        print("Arguments: (* is used to indicate that argument is not required)")
        print("  certificate_jpg_file : String (File Path for Certificate in .jpg format)")
        print("  csv_file : String (File path for recipients in .csv format)")
        print("  text_pos_tuple : Tuple (Split by \',\')")
        print("  *font_color_RGB_tuple : Tuple (Split by \',\') , default = (0,0,0)")
        print("  *text_alignment : String , default = center")
        print("  *font_path : String , default = None")
        print("  *font_size : Integer , default = 30")
        sys.exit(0)
    elif len(args) == 0:
        print("Error: Arguments Required!")
        print("Please type \"aot.py -h\" for help.")
        sys.exit(0)
    for arg in args:
        if "=" in arg:
            print(arg)
            key,value = arg.split("=")
            if key == "certificate_jpg_file":
                certificate_jpg_file = str(value)
            elif key == "csv_file":
                csv_file = str(value)
            elif key == "text_pos_tuple":
                text_pos_tuple = tuple(value)
            elif key == "text_anchor":
                text_anchor = str(value)
            elif key == "font_color_RGB_tuple":
                font_color_RGB_tuple = tuple(value)
            elif key == "text_alignment":
                text_alignment = str(value)
            elif key == "font_path":
                font_path = str(value)
            elif key == "font_size":
                font_size = int(value)
            else:
                print(f"Unknown argument '{key}'")
        else:
            print(arg)
            try:
                if certificate_jpg_file == None:
                    certificate_jpg_file = str(arg)
                elif csv_file == None:
                    csv_file = str(arg)
                elif text_pos_tuple == None:
                    text_pos_tuple = tuple(arg)    
                else:
                    print("Extra argument provided, but all default values are already replaced.")
                    break
            except ValueError:
                print(f"Invalid argument '{arg}'")
    
    email = input("Please enter your email address:\n")
    app_password = input("Please input your app password (This is required to send the emails):\n")
    
            
    
