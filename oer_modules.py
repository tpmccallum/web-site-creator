import os
import xlrd
import urllib
import libxml2
from BeautifulSoup import BeautifulSoup


def initialize():
    print "Initializing"
    htmlDirectory = os.path.join(os.getcwd(), 'html')
    if not os.path.exists(htmlDirectory):
    	os.makedirs(htmlDirectory)
    	return htmlDirectory
    else:
    	print "Directory %s already exists, please remove and try again" % (htmlDirectory)
    	exit()

def openSpreadsheetFile():
    spreadsheetLocation = os.path.join(os.getcwd(), "course_template.xlsx")
    print "Opening spreadsheet located at %s " % (spreadsheetLocation)
    spreadsheet = xlrd.open_workbook(spreadsheetLocation)
    print "Spreadsheet opened successfully"
    return spreadsheet

def getStyleConfiguration(firstSheet):
    headerLogoLeft = firstSheet.cell(0,1).value
    print "Header Logo Left %s " % (headerLogoLeft)
    headerLogoRight = firstSheet.cell(1,1).value
    print "Header Logo Right %s " % (headerLogoRight)
    ColourDark = firstSheet.cell(2,1).value
    print "Colour Dark %s " % (ColourDark)
    ColourLight = firstSheet.cell(3,1).value
    print "Colour Light %s " % (ColourLight)
    CourseCode = firstSheet.cell(4,1).value
    print "Course Code %s " % (CourseCode)
    CourseName = firstSheet.cell(5,1).value
    print "Course Name %s " % (CourseName)
    MicroCode = firstSheet.cell(6,1).value
    print "Micro Code %s " % (MicroCode)
    MicroName = firstSheet.cell(7,1).value
    print "Micro Name %s " % (MicroName)
    GradientFill = "( " + ColourDark + ", " + ColourLight + " )"
    styleConfiguration = {'headerLogoLeft' : headerLogoLeft, 'headerLogoRight' : headerLogoRight, 'ColourDark' : ColourDark, 'ColourLight' : ColourLight, 'CourseCode' : CourseCode, 'CourseName' : CourseName, 'MicroCode' : MicroCode, 'MicroName' : MicroName, 'GradientFill' : GradientFill}
    return styleConfiguration

def getContent(secondSheet):
	print "Creating list"
	l = list()
	print "Creating dictionary"
	d = dict()
	print "Setting num to zero"
	num = 0
	print "Iterating through worksheet"
	for row in secondSheet.col(1):
		num = num + 1 
		print "num = %d" % (num)
		l.append(row.value)
		if num == 3:
			print "num equals 3"
			print "Adding list to dictionary"
			d[row.value] = l
			print "Re-setting num to zero"
			num = 0
			print "Emptying the list"
			l = list()
	return d

def createCSS(htmlDirectory, styleConfiguration):
	print "HTML Directory %s " % (htmlDirectory)
	file2 = open(os.path.join(htmlDirectory, 'custom_oer.css'), 'w')
	ContainerString = """.container{
	background-color:#efe9e5;
	background: -webkit-linear-gradient""" + styleConfiguration.get('GradientFill') + """!important;
	background: -o-linear-gradient""" + styleConfiguration.get('GradientFill') + """!important;
	background: -moz-linear-gradient""" + styleConfiguration.get('GradientFill') + """!important;
	background: linear-gradient""" + styleConfiguration.get('GradientFill') + """!important;
	border:1px solid #5b5b5b;
	border-top-width: 0px;
	-ms-box-shadow: 1px 4px 4px 0px #1E1E1E !important;
	-moz-box-shadow: 1px 4px 4px 0px #1E1E1E !important;
	-webkit-box-shadow: 1px 4px 4px 0px #1E1E1E !important;
	box-shadow: 1px 4px 4px 0px #1E1E1E !important;
	}"""
	HeaderString = """.header-oer{ 
	background: #FFF!important; 
	text-align: center!important; 
	border:1px solid #5b5b5b!important; 
	border-top-width: 0px!important;
	-ms-box-shadow: 3px 4px 6px 0px #1E1E1E !important;
    -moz-box-shadow: 3px 4px 6px 0px #1E1E1E !important;
    -webkit-box-shadow: 3px 4px 6px 0px #1E1E1E !important;
    box-shadow: 3px 4px 6px 0px #1E1E1E !important;
    }"""
	FooterString = """.footer-wrap {
	background-color: #000;
	min-height: 70px;
	color: #fff;
	padding-top: 20px;
	font-size: 12px;
	padding: 20px 10px;
	text-align: center;
	border-top-width: 2px!important;
	-ms-box-shadow: 3px 4px 6px 0px #1E1E1E !important;
	-moz-box-shadow: 3px 4px 6px 0px #1E1E1E !important;
	-webkit-box-shadow: 3px 4px 6px 0px #1E1E1E !important;
	box-shadow: 3px 4px 6px 0px #1E1E1E !important;
	}
	.footer-wrap p {
	line-height: 18px;
	text-align: center;
	padding-bottom: 10px;
	}
	.footer-wrap .right {
	text-align: center;
	padding-bottom: 20px;
	}
	.footer-wrap a {
	color: white;
	}
	@media (min-width: 766px) {
	.footer-wrap {
	text-align: center;
	padding-bottom: 0px;
	}
	.footer-wrap .right {
	float: right;
	text-align: center;
	padding-right: 10px;
	padding-left: 20px;
	}
	.footer-wrap p {
	line-height: 15px;
	text-align: center;
	padding-bottom: 0px;
	margin-bottom: 2px;
	}
	}"""
	OerImageString = """.image-oer {
    position: relative;
	display: block;
	max-width: 100%;
	left: 0;
	right: 0;
	bottom: 0;
	margin: auto;
	}"""
	OerGreen = """.green-oer > li > a {
	color : green!important;
	}"""
	file2.write(OerGreen)
	file2.write(OerImageString)
	file2.write(ContainerString)
	file2.write(HeaderString)
	file2.write(FooterString)
	print OerGreen
	print OerImageString
	print ContainerString
	print HeaderString
	print FooterString

def fetchWikicontent(wikiUrl):
	response = urllib.urlopen(wikiUrl)
	html = response.read()
	soup = BeautifulSoup(html)
	cleanHtml = soup.prettify()
	doc = libxml2.parseDoc(cleanHtml)
	ctxt = doc.xpathNewContext()
	res = ctxt.xpathEval("//*[@id='bodyContent']/p")
	return res


def createHtml(htmlDirectory, styleConfiguration, wikiContent, mediaUrl, pageName):
	print "HTML Directory %s " % (htmlDirectory)
	file = open(os.path.join(htmlDirectory, pageName + '.html'), 'w')
	#print type(wikiContent)
	#print wikiContent
	HtmlString = """<!DOCTYPE html>
	<html lang="en">
	<head>
	<meta charset="utf-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<meta name="description" content="">
	<meta name="author" content="">
	<link rel="icon" href="../../favicon.ico">

	<title>""" + pageName + """</title>

	<!-- Bootstrap core CSS -->
	<!-- Latest compiled and minified CSS -->
	<link rel="stylesheet" href="custom_oer.css">
	<link rel="stylesheet" href="http://maxcdn.bootstrapcdn.com/bootstrap/3.2.0/css/bootstrap.min.css">


	<!-- Optional theme -->
	<link rel="stylesheet" href="http://maxcdn.bootstrapcdn.com/bootstrap/3.2.0/css/bootstrap-theme.min.css">

	<!-- Latest compiled and minified JavaScript -->
	<script src="http://maxcdn.bootstrapcdn.com/bootstrap/3.2.0/js/bootstrap.min.js"></script>

	<!-- Custom styles for this template -->
	<link href="http://getbootstrap.com/examples/justified-nav/justified-nav.css" rel="stylesheet">

	<!-- Just for debugging purposes. Don't actually copy these 2 lines! -->
	<!--[if lt IE 9]><script src="http://getbootstrap.com/assets/js/ie-emulation-modes-warning.js"></script><![endif]-->
	<script src="http://getbootstrap.com/assets/js/ie-emulation-modes-warning.js"></script>

	<!-- IE10 viewport hack for Surface/desktop Windows 8 bug -->
	<script src="http://getbootstrap.com/assets/js/ie10-viewport-bug-workaround.js"></script>

	<!-- HTML5 shim and Respond.js IE8 support of HTML5 elements and media queries -->
	<!--[if lt IE 9]>
	<script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
	<script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
	<![endif]-->
	</head>

	<body>
	<div class="row header-oer">
	<div class="col-md-4"><img src=""" + styleConfiguration.get('headerLogoLeft') + """ alt="OERu"/></div>
	<div class="col-md-4"><h2>""" + styleConfiguration.get('CourseCode') + """ </h2> <br /><b> """ + styleConfiguration.get('CourseName') + """</b></div>
	<div class="col-md-4"><img src='""" + styleConfiguration.get('headerLogoRight') + """' alt="OERu"/></div>
	</div>  

	<div class="container">	
	<div class="masthead">
	<ul class="nav nav-justified">
	<li class="active"><a href="#"> """ + pageName + """</a></li>
	<li><a href="#">Projects</a></li>
	<li><a href="#">Services</a></li>
	<li><a href="#">Downloads</a></li>
	<li><a href="#">About</a></li>
	<li><a href="#">Contact</a></li>
	</ul>
	</div>
	<br />

	<!-- Example row of columns -->
	<div class="row">
	<div class="col-md-8">
	<h2>""" + pageName + """</h2>
	<p> """ + str(wikiContent) + """</p>
	</div>
	<div class="col-md-4">
	<img class='image-oer' src=' """ + mediaUrl + """ '</div>	
	</div>
	<br />
	</div>

	
	<br />


	</div> <!-- /container -->
	<!-- Site footer -->
	<div class="row">
	<div class="col-md-1"></div>
	<div class="col-md-10">
	<div class="footer-wrap" style="text-align: center;">	
	<p>#TODO ABN: 40 234 732 081 | CRICOS: QLD 00244B<span class="line-break">, </span>
	NSW 02225M | TEQSA: PRV12081<span class="line-break"> | </span>
	<a href="#">Disclaimer</a><span class="line-break"> | <a href="#">Contact us</a></p>
	<p>University of Southern Queensland<span class="line-break"></span> #TODO</p>
	</div>
	</div>
	<div class="col-md-1"></div>

	</div>


	<!-- Bootstrap core JavaScript
	================================================== -->
	<!-- Placed at the end of the document so the pages load faster -->
	</body>
	</html>"""
	file.write(HtmlString)
	file.close()
