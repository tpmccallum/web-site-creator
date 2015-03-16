import oer_modules
htmlDirectory = oer_modules.initialize()
spreadsheet = oer_modules.openSpreadsheetFile()
firstSheet = spreadsheet.sheet_by_index(0)
secondSheet = spreadsheet.sheet_by_index(1)
styleConfiguration = oer_modules.getStyleConfiguration(firstSheet)
contentDict = oer_modules.getContent(secondSheet)
oer_modules.createCSS(htmlDirectory, styleConfiguration)
for k, v in contentDict.items():
	wikiUrl = v[0]
	print "wikiUrl is %s " % (wikiUrl)
	wikiContent = oer_modules.fetchWikicontent(wikiUrl)
	mediaUrl = v[1]
	print "mediaUrl is %s " % (mediaUrl)
	pageName = v[2]
	print "pageName is %s " % (pageName)
	oer_modules.createHtml(htmlDirectory, styleConfiguration, wikiContent, mediaUrl, pageName)
