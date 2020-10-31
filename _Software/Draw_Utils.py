#write pdf
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.units import inch
width, height = A4 #useful for drawing functions
from PIL import Image
#other utils
from Misc_Utils import string_split
#currency 
import locale
locale.setlocale(locale.LC_ALL, '')

#column constants for x positioning
SERVICE = .8*inch
QTE = 4.75*inch  
MONTANT = 5.4*inch
TAUX = 6.25*inch
COMMISSION = 6.75*inch
linespace = 15
RIGHT = 25

#color constants for Celestial Blue #4992C9
R = 73/256
G = 146/256
B = 201/256



#Canvas drawing functions. Template is called for each new invoice and others are called only if appropriate for a given therapist
def draw_invoice_template(c, name, year, month, day):
	current_pos = height - (0.5*inch)
	c.setFont('Helvetica-Bold',10)
	c.drawString(0.5 * inch, current_pos ,"Divan Bleu")
	c.drawImage("_Software\logo.png", width - 3*inch, current_pos - 1.2*inch, width=2*inch, height=1.5*inch, mask=[200,255,200,255,200,255]) # (img, x,y, width, height, mask)
	current_pos -= linespace 
	
	c.setFont('Helvetica',10)
	c.drawString(0.5*inch,current_pos, "1150 Boulevard Saint-Joseph Est #108")
	current_pos -= linespace
	uniline = u"Montr\u00e9al QC H2J 1L5" #Montreal
	c.drawString(0.5*inch,current_pos, uniline.strip())
	current_pos -= linespace
	c.drawString(0.5*inch,current_pos, "info@divanbleu.com")
	current_pos -= linespace
	c.drawString(0.5*inch,current_pos, "www.divanbleu.com")
	current_pos -= linespace
	uniline = u"N. d'enregistrement de la TPS/TVH : TPS" # TODO Want unicode character \u2116 here instead of N.
	c.drawString(0.5*inch,current_pos, uniline.strip()) 
	current_pos -= linespace
	c.drawString(0.5*inch,current_pos, "(783019672) RT 0001")
	current_pos -= linespace
	uniline = u"N. d'enregistrement de la TVQ : TVQ" # TODO Same unicode character here
	c.drawString(0.5*inch,current_pos, uniline)
	current_pos -= linespace
	c.drawString(0.5*inch,current_pos, "(1226687561) TQ0001")
	current_pos -= (linespace + 10)

	c.setFont('Helvetica', 20)
	c.setFillColorRGB(R,G,B) 
	c.drawString(0.5*inch,current_pos,"FACTURE")
	current_pos -= 30

	pos2 = current_pos
	c.setFont('Helvetica-Bold', 10)
	c.setFillColorRGB(0,0,0)
	#left side
	uniline = u"FACTURER \u00c0" #FACTURER A
	c.drawString(0.5*inch,current_pos, uniline.strip())
	#right bold
	c.drawRightString(width - 2*inch, current_pos, "DATE")
	current_pos -= linespace
	#left
	c.setFont('Helvetica', 14)
	c.drawString(0.5*inch,current_pos,name)
	
	c.setFont('Helvetica-Bold', 10)
	uniline = u"MODALIT\u00C9S" #MODALITES
	c.drawRightString(width - 2*inch, current_pos, uniline.strip())
	current_pos -= 30

	#right non-bold
	c.setFont('Helvetica', 10)
	datestring = "{}-{}-{}".format(year,month,day)
	c.drawString(width - 1.85*inch, pos2, datestring)
	pos2 -= linespace
	uniline = u"Payable d\u00e8s r\u00e9ception"
	c.drawString(width - 1.85*inch, pos2, uniline.strip())

	#start of mid section
	c.setFillColorRGB(R,G,B)  
	c.line(.5*inch, current_pos, width - .5*inch, current_pos)
	current_pos -= 4*linespace
	c.rect(.5*inch, current_pos, width - inch, 1.5*linespace, stroke=False, fill=True)
	current_pos += .5*linespace
	

	return current_pos #return the position for the next draw functions to keep drawing from current position


def draw_service(c, position, service, rate, revenue, commission, quantity):
	current_pos = position
	c.setFillColorRGB(1,1,1) 
	c.drawString(SERVICE,current_pos,"SERVICE")
	uniline = u"QT\u00C9"
	c.drawString(QTE, current_pos, uniline.strip())
	
	c.drawString(MONTANT, current_pos, "MONTANT")
	c.drawString(TAUX, current_pos, "TAUX")
	c.drawString(COMMISSION, current_pos, "COMMISSION")
	current_pos -= 1.5*linespace
	service_lines = string_split(service, length=60)
	c.setFont('Helvetica', 10)
	c.setFillColorRGB(0,0,0)
	c.drawRightString(QTE + RIGHT, current_pos, str(quantity))
	c.drawRightString(MONTANT + 2*RIGHT, current_pos, locale.currency(revenue, grouping =True))
	c.drawRightString(TAUX + RIGHT, current_pos, rate)
	c.drawRightString(COMMISSION + 2*RIGHT, current_pos, locale.currency(commission, grouping=True))
	for line in service_lines:
		c.drawString(SERVICE, current_pos, line)
		current_pos  -= linespace
		if current_pos < 135:
			c.showPage()
			current_pos = height - (0.5*inch)
	current_pos -= 0.5*linespace
	return current_pos

def draw_rent(c, position, rent):
	current_pos = position
	c.setFillColorRGB(1,1,1) 
	c.drawString(SERVICE,current_pos,"SERVICE")
	c.drawString(COMMISSION, current_pos, "MONTANT")
	current_pos -= 1.5*linespace
	c.setFont('Helvetica', 10)
	c.setFillColorRGB(0,0,0)
	uniline = u"Loyer - Bureau de th\u00e9rapie"
	c.drawString(SERVICE, current_pos, uniline.strip())
	c.drawRightString(COMMISSION + 2*RIGHT, current_pos, locale.currency(rent, grouping=True))
	current_pos -= 1.5*linespace
	return current_pos

def draw_end(c, position, partial, TPS, TVQ, total, pays_tax):
	current_pos = position
	c.setFillColorRGB(0,0,0)
	c.setDash(6,3) 
	c.line(inch, current_pos, width - inch, current_pos)
	current_pos -= 3*linespace
	pos2 = current_pos

	#right side totals
	c.setFont('Helvetica', 10)
	c.drawRightString(COMMISSION + 2*RIGHT, current_pos, locale.currency(partial, grouping=True))
	current_pos -= linespace
	if pays_tax:
		c.drawRightString(COMMISSION+ 2*RIGHT, current_pos, locale.currency(TPS, grouping=True))
		current_pos -= linespace
		c.drawRightString(COMMISSION + 2*RIGHT, current_pos, locale.currency(TVQ, grouping=True))
		current_pos -= linespace
	current_pos -= 0.5*linespace
	c.setFont("Helvetica-Bold", 11)
	c.drawRightString(COMMISSION + 2*RIGHT, current_pos, locale.currency(total, grouping=True))
	
	#left side titles
	c.setFont('Helvetica', 10)
	c.drawString(QTE, pos2, "TOTAL PARTIEL")
	pos2 -= linespace
	if pays_tax:
		c.drawString(QTE, pos2, "TPS @ 5%")
		pos2 -= linespace
		c.drawString(QTE, pos2, "TVQ @ 9,975%")
		pos2 -= linespace
	c.setFont("Helvetica-Bold", 11)
	c.drawString(QTE, pos2, "TOTAL")

	current_pos -= 3*linespace
	c.setFillColorRGB(R,G,B)
	c.rect(.5*inch, current_pos, width - inch, 1.5*linespace, stroke=False, fill=True)
	


