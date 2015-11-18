fs = require('fs')
DocxGen = require('docxtemplater')
expect = require('chai').expect

fileNames = [
	'chartExample.docx',
	'loopChartExample.docx',
	'dateExample.docx',
	'multipleChartsExample.docx'
]

ChartModule = require('../js/index.js')

docX = {}

loadFile = (name) ->
	if fs.readFileSync? then return fs.readFileSync(__dirname + "/../examples/" + name, "binary")
	xhrDoc = new XMLHttpRequest()
	xhrDoc.open('GET', "../examples/" + name, false)
	if (xhrDoc.overrideMimeType)
		xhrDoc.overrideMimeType('text/plain; charset=x-user-defined')
	xhrDoc.send()
	xhrDoc.response

for name in fileNames
	content=loadFile(name)
	docX[name]=new DocxGen()
	docX[name].loadedContent=content

describe 'adding with {$ chart} syntax', () ->
	name = 'chartExample.docx'
	chartModule = new ChartModule()
	docX[name].attachModule(chartModule)
	out = docX[name]
		.load(docX[name].loadedContent)
		.setData({
			chart: {
				options: {
					width: 300,
					height: 200,
					border: false,
					legend: {
						position: 'l' # can be 'r'
					}
				},
				lines: [
					{
						name: 'line 1',
						data: [
							{
								x: 'day 1',
								y: '4.3'
							},
							{
								x: 'day 2',
								y: '2.5'
							},
							{
								x: 'day 3',
								y: '3.5'
							},
							{
								x: 'day 4',
								y: '4.5'
							}
						]
					},
					{
						name: 'line 2',
						data: [
							{
								x: 'day 1',
								y: '2.4'
							},
							{
								x: 'day 2',
								y: '4.4000000000000004'
							},
							{
								x: 'day 3',
								y: '1.8'
							},
							{
								x: 'day 4',
								y: '2.8'
							}
						]
					},
					{
						name: 'line 3',
						data: [
							{
								x: 'day 1',
								y: '2'
							},
							{
								x: 'day 2',
								y: '2'
							},
							{
								x: 'day 3',
								y: '3'
							},
							{
								x: 'day 4',
								y: '5'
							}
						]
					}
				]
			}
		})
		.render()
	zip = out.getZip()

	it 'should create relationship in rels file', () ->
		relsFile = zip.files['word/_rels/document.xml.rels']
		expect(relsFile?).to.equal(true)
		relsFileContent = relsFile.asText()
		expect(relsFileContent).to.equal("""
			<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart6.xml"/></Relationships>
		""")

	it 'should create chart file without min and max in scaling', () ->
		chartFile = zip.files['word/charts/chart6.xml']
		expect(chartFile?).to.equal(true)
		chartFileContent = chartFile.asText()
		expect(chartFileContent).to.contain("c:orientation")
		expect(chartFileContent).to.not.contain("c:max")


	it 'should add content type', () ->
		typeFile = zip.files['[Content_Types].xml']
		expect(typeFile?).to.equal(true)
		typeFileContent = typeFile.asText()
		expect(typeFileContent).to.equal("""
			<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="wmf" ContentType="image/x-wmf"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/><Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml" PartName="/word/charts/chart6.xml"/></Types>
		""")





	fs.writeFile('test.docx', zip.generate({type:"nodebuffer"}));

describe 'adding with {$ chart} syntax inside a loop', () ->
	name = 'loopChartExample.docx'
	chartModule = new ChartModule()
	docX[name].attachModule(chartModule)
	out = docX[name]
	.load(docX[name].loadedContent)
	.setData({
			subsidiaries: [
				{
					title: "Euro Giant Ltd",
					chart: {
						options: {
							width: 300,
							height: 200,
							border: false,
							legend: {
								position: 'l' # can be 'r'
							}
						},
						lines: [
							{
								name: 'Product A',
								data: [
									{
										x: 'Jan',
										y: '14.5'
									},
									{
										x: 'Feb',
										y: '12.5'
									},
									{
										x: 'Mar',
										y: '13.5'
									}
								]
							},
							{
								name: 'Product B',
								data: [
									{
										x: 'Jan',
										y: '32.4'
									},
									{
										x: 'Feb',
										y: '34.4000000000000004'
									},
									{
										x: 'Mar',
										y: '31.8'
									}
								]
							},
							{
								name: 'Product C',
								data: [
									{
										x: 'Jan',
										y: '32'
									},
									{
										x: 'Feb',
										y: '22'
									},
									{
										x: 'Mar',
										y: '35'
									}
								]
							}
						]
					}
				},
				{
					title: "USA Giant Inc",
					chart: {
						options: {
							width: 300,
							height: 200,
							border: false,
							legend: {
								position: 'l' # can be 'r'
							}
						},
						lines: [
							{
								name: 'Product A',
								data: [
									{
										x: 'Jan',
										y: '15'
									},
									{
										x: 'Feb',
										y: '25'
									},
									{
										x: 'Mar',
										y: '12'
									}
								]
							},
							{
								name: 'Product B',
								data: [
									{
										x: 'Jan',
										y: '7'
									},
									{
										x: 'Feb',
										y: '8.5'
									},
									{
										x: 'Mar',
										y: '27'
									}
								]
							},
							{
								name: 'Product C',
								data: [
									{
										x: 'Jan',
										y: '20'
									},
									{
										x: 'Feb',
										y: '18'
									},
									{
										x: 'Mar',
										y: '22'
									}
								]
							}
						]
					}
				}
			]

		})
	.render()
	zip = out.getZip()

	it 'should create two relationships in rels file', () ->
		relsFile = zip.files['word/_rels/document.xml.rels']
		expect(relsFile?).to.equal(true)
		relsFileContent = relsFile.asText()
		expect(relsFileContent).to.equal("""
			<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects" Target="stylesWithEffects.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart7.xml"/><Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart8.xml"/></Relationships>
		""")

	it 'should create two separate chart files', () ->
		chartFile1 = zip.files['word/charts/chart7.xml']
		expect(chartFile1?).to.equal(true)
		chartFile2 = zip.files['word/charts/chart8.xml']
		expect(chartFile2?).to.equal(true)

	it 'should add content type', () ->
		typeFile = zip.files['[Content_Types].xml']
		expect(typeFile?).to.equal(true)
		typeFileContent = typeFile.asText()
		expect(typeFileContent).to.equal("""
			<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/stylesWithEffects.xml" ContentType="application/vnd.ms-word.stylesWithEffects+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml" PartName="/word/charts/chart7.xml"/><Override ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml" PartName="/word/charts/chart8.xml"/></Types>
		""")





	fs.writeFile('looptest.docx', zip.generate({type:"nodebuffer"}));

describe 'multiple charts adding', () ->
	name = 'multipleChartsExample.docx'
	chartModule = new ChartModule()
	docX[name].attachModule(chartModule)
	out = docX[name]
		.load(docX[name].loadedContent)
		.setData({
			chart: {
				lines: [
					{
						name: 'line 1',
						data: [
							{
								x: 'day 1',
								y: '4.3'
							},
							{
								x: 'day 2',
								y: '2.5'
							},
							{
								x: 'day 3',
								y: '3.5'
							},
							{
								x: 'day 4',
								y: '4.5'
							}
						]
					},
					{
						name: 'line 2',
						data: [
							{
								x: 'day 1',
								y: '2.4'
							},
							{
								x: 'day 2',
								y: '4'
							},
							{
								x: 'day 3',
								y: '1.8'
							},
							{
								x: 'day 4',
								y: '2.8'
							}
						]
					},
					{
						name: 'line 3',
						data: [
							{
								x: 'day 1',
								y: '2'
							},
							{
								x: 'day 2',
								y: '2'
							},
							{
								x: 'day 3',
								y: '3'
							},
							{
								x: 'day 4',
								y: '5'
							}
						]
					}
				]
			},
			chart2: {
				options: {
					legend: {
						position: 'b'
					}
					axis: {
						y: {
							min: 500,
							max: 5000
						}
					}
				}
				lines: [
					{
						name: 'line 1',
						data: [
							{
								x: 'day 1',
								y: '1000'
							},
							{
								x: 'day 2',
								y: '1000'
							},
							{
								x: 'day 3',
								y: '3000'
							},
							{
								x: 'day 4',
								y: '4500'
							},
							{
								x: 'day 5',
								y: '1000'
							},
							{
								x: 'day 6',
								y: '3000'
							},
							{
								x: 'day 7',
								y: '1000'
							},
							{
								x: 'day 8',
								y: '3000'
							},
							{
								x: 'day 9',
								y: '1000'
							},
							{
								x: 'day 10',
								y: '3000'
							},
							{
								x: 'day 11',
								y: '1000'
							},
							{
								x: 'day 12',
								y: '3000'
							},
							{
								x: 'day 13',
								y: '1000'
							},
							{
								x: 'day 14',
								y: '3000'
							},
							{
								x: 'day 15',
								y: '1000'
							},
							{
								x: 'day 16',
								y: '3000'
							},
							{
								x: 'day 17',
								y: '1000'
							},
							{
								x: 'day 18',
								y: '3000'
							},
							{
								x: 'day 19',
								y: '1000'
							},
							{
								x: 'day 20',
								y: '3000'
							},
							{
								x: 'day 21',
								y: '1000'
							},
							{
								x: 'day 22',
								y: '3000'
							},
							{
								x: 'day 23',
								y: '1000'
							},
							{
								x: 'day 24',
								y: '3000'
							},
							{
								x: 'day 25',
								y: '1000'
							},
							{
								x: 'day 26',
								y: '3000'
							},
							{
								x: 'day 37',
								y: '1000'
							},
							{
								x: 'day 38',
								y: '3000'
							},
							{
								x: 'day 39',
								y: '1000'
							},
							{
								x: 'day 30',
								y: '3000'
							},
							{
								x: 'day 31',
								y: '1000'
							},
							{
								x: 'day 32',
								y: '3000'
							},
							{
								x: 'day 33',
								y: '1000'
							},
							{
								x: 'day 34',
								y: '3000'
							},
							{
								x: 'day 35',
								y: '1000'
							},
							{
								x: 'day 36',
								y: '3000'
							},
							{
								x: 'day 37',
								y: '1000'
							},
							{
								x: 'day 38',
								y: '3000'
							},
							{
								x: 'day 39',
								y: '1000'
							},
							{
								x: 'day 40',
								y: '3000'
							},
							{
								x: 'day 41',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 43',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 45',
								y: '1000'
							},
							{
								x: 'day 46',
								y: '3000'
							},
							{
								x: 'day 38',
								y: '3000'
							},
							{
								x: 'day 39',
								y: '1000'
							},
							{
								x: 'day 40',
								y: '3000'
							},
							{
								x: 'day 41',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 43',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 45',
								y: '1000'
							},
							{
								x: 'day 46',
								y: '3000'
							},
							{
								x: 'day 38',
								y: '3000'
							},
							{
								x: 'day 39',
								y: '1000'
							},
							{
								x: 'day 40',
								y: '3000'
							},
							{
								x: 'day 41',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 43',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 45',
								y: '1000'
							},
							{
								x: 'day 46',
								y: '3000'
							},
							{
								x: 'day 40',
								y: '3000'
							},
							{
								x: 'day 41',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 43',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 45',
								y: '1000'
							},
							{
								x: 'day 40',
								y: '3000'
							},
							{
								x: 'day 41',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 43',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 45',
								y: '1000'
							},
							{
								x: 'day 40',
								y: '3000'
							},
							{
								x: 'day 41',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 43',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 45',
								y: '1000'
							},
							{
								x: 'day 40',
								y: '3000'
							},
							{
								x: 'day 41',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 43',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 45',
								y: '1000'
							},
							{
								x: 'day 40',
								y: '3000'
							},
							{
								x: 'day 41',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 43',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 45',
								y: '1000'
							},
							{
								x: 'day 40',
								y: '3000'
							},
							{
								x: 'day 41',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 43',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 45',
								y: '1000'
							},
							{
								x: 'day 40',
								y: '3000'
							},
							{
								x: 'day 41',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 43',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 45',
								y: '1000'
							},
							{
								x: 'day 40',
								y: '3000'
							},
							{
								x: 'day 41',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 43',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 45',
								y: '1000'
							},
							{
								x: 'day 40',
								y: '3000'
							},
							{
								x: 'day 41',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 43',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 45',
								y: '1000'
							},
							{
								x: 'day 40',
								y: '3000'
							},
							{
								x: 'day 41',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 43',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 45',
								y: '1000'
							},
							{
								x: 'day 40',
								y: '3000'
							},
							{
								x: 'day 41',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 43',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 45',
								y: '1000'
							},
							{
								x: 'day 40',
								y: '3000'
							},
							{
								x: 'day 41',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 43',
								y: '1000'
							},
							{
								x: 'day 44',
								y: '3000'
							},
							{
								x: 'day 45',
								y: '1000'
							}
						]
					}
				]
			}
		})
		.render()
	zip = out.getZip()

	fs.writeFile('multipleTest.docx', zip.generate({type: "nodebuffer"}));

describe 'date type with unix timestapm', () ->
	name = 'dateExample.docx'
	chartModule = new ChartModule()
	docX[name].attachModule(chartModule)
	out = docX[name]
		.load(docX[name].loadedContent)
		.setData({
			chart: {
				options: {
					axis: {
						x: {
							type: 'date', # can be 'date', 'normal'
							date: {
								format: 'unix',
								code: 'm/yyyy',
								unit: 'days',
								step: '2'
							}
						}
					}
				},
				lines: [
					{
						name: 'line 1',
						data: [
							{
								x: '1382400000',
								y: 2
							},
							{
								x: '1382486400',
								y: 4
							},
							{
								x: '1382572800',
								y: 3
							},
							{
								x: '1382659200',
								y: 5
							}
						]
					}
				]
			}
		})
		.render()
	zip = out.getZip()
	fs.writeFile('date.docx', zip.generate({type:"nodebuffer"}));

	# it 'should work with one image',()->
	# 	name='imageExample.docx'
	# 	imageModule=new ImageModule({centered:false})
	# 	docX[name].attachModule(imageModule)
	# 	out=docX[name]
	# 		.load(docX[name].loadedContent)
	# 		.setData({image:'examples/image.png'})
	# 		.render()

	# 	zip=out.getZip()

	# 	imageFile=zip.files['word/media/image_generated_1.png']
	# 	expect(imageFile?).to.equal(true)
	# 	expect(imageFile.asText().length).to.equal(17417)

	# 	relsFile=zip.files['word/_rels/document.xml.rels']
	# 	expect(relsFile?).to.equal(true)
	# 	relsFileContent=relsFile.asText()
	# 	expect(relsFileContent).to.equal("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="hId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header0.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image_generated_1.png"/></Relationships>""")

	# 	documentFile=zip.files['word/document.xml']
	# 	expect(documentFile?).to.equal(true)
	# 	documentContent=documentFile.asText()
	# 	# expect(documentContent).to.equal("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr></w:pPr><w:r><w:rPr></w:rPr><w:drawing>\n  <wp:inline distT="0" distB="0" distL="0" distR="0">\n    <wp:extent cx="1905000" cy="1905000"/>\n    <wp:effectExtent l="0" t="0" r="0" b="0"/>\n    <wp:docPr id="2" name="Image 2" descr="description"/>\n    <wp:cNvGraphicFramePr>\n      <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>\n    </wp:cNvGraphicFramePr>\n    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">\n        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">\n          <pic:nvPicPr>\n            <pic:cNvPr id="0" name="Picture 1" descr="description"/>\n            <pic:cNvPicPr>\n              <a:picLocks noChangeAspect="1" noChangeArrowheads="1"/>\n            </pic:cNvPicPr>\n          </pic:nvPicPr>\n          <pic:blipFill>\n            <a:blip r:embed="rId6">\n              <a:extLst>\n                <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">\n                  <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>\n                </a:ext>\n              </a:extLst>\n            </a:blip>\n            <a:srcRect/>\n            <a:stretch>\n              <a:fillRect/>\n            </a:stretch>\n          </pic:blipFill>\n          <pic:spPr bwMode="auto">\n            <a:xfrm>\n              <a:off x="0" y="0"/>\n              <a:ext cx="1905000" cy="1905000"/>\n            </a:xfrm>\n            <a:prstGeom prst="rect">\n              <a:avLst/>\n            </a:prstGeom>\n            <a:noFill/>\n            <a:ln>\n              <a:noFill/>\n            </a:ln>\n          </pic:spPr>\n        </pic:pic>\n      </a:graphicData>\n    </a:graphic>\n  </wp:inline>\n</w:drawing></w:r><w:bookmarkStart w:id="13" w:name="_GoBack"/><w:bookmarkEnd w:id="13"/></w:p><w:sectPr><w:headerReference w:type="default" r:id="hId0"/><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/><w:pgMar w:top="2810" w:left="1800" w:right="1800" w:bottom="1440"/><w:cols w:num="1" w:sep="off" w:equalWidth="1"/></w:sectPr></w:body></w:document>""")

	# 	fs.writeFile("test.docx",zip.generate({type:"nodebuffer"}));

	# it 'should work with centering',()->
	# 	d=new DocxGen()
	# 	name='imageExample.docx'
	# 	imageModule=new ImageModule({centered:true})
	# 	d.attachModule(imageModule)
	# 	out=d
	# 		.load(docX[name].loadedContent)
	# 		.setData({image:'examples/image.png'})
	# 		.render()

	# 	zip=out.getZip()
	# 	imageFile=zip.files['word/media/image_generated_1.png']
	# 	expect(imageFile?).to.equal(true)
	# 	expect(imageFile.asText().length).to.equal(17417)

	# 	relsFile=zip.files['word/_rels/document.xml.rels']
	# 	expect(relsFile?).to.equal(true)
	# 	relsFileContent=relsFile.asText()
	# 	expect(relsFileContent).to.equal("""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\" Target=\"numbering.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\" Target=\"footnotes.xml\"/><Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes\" Target=\"endnotes.xml\"/><Relationship Id=\"hId0\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header0.xml\"/><Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image_generated_1.png\"/></Relationships>""")

	# 	documentFile=zip.files['word/document.xml']
	# 	expect(documentFile?).to.equal(true)
	# 	documentContent=documentFile.asText()
	# 	# expect(documentContent).to.equal("""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wx=\"http://schemas.microsoft.com/office/word/2003/auxHint\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>\n\t<w:pPr>\n\t<w:jc w:val=\"center\"/>\n  </w:pPr>\n  <w:r>\n\t<w:rPr/>\n\t<w:drawing>\n\t  <wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">\n\t\t<wp:extent cx=\"1905000\" cy=\"1905000\"/>\n\t\t<wp:docPr id=\"15\" name=\"rId6.png\"/>\n\t\t<a:graphic>\n\t\t  <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\n\t\t\t<pic:pic>\n\t\t\t  <pic:nvPicPr>\n\t\t\t\t<pic:cNvPr id=\"15\" name=\"rId6.png\"/>\n\t\t\t\t<pic:cNvPicPr/>\n\t\t\t  </pic:nvPicPr>\n\t\t\t  <pic:blipFill>\n\t\t\t\t<a:blip r:embed=\"rId6\"/>\n\t\t\t\t</pic:blipFill>\n\t\t\t  <pic:spPr>\n\t\t\t\t<a:xfrm>\n\t\t\t\t  <a:off x=\"0\" y=\"0\"/>\n\t\t\t\t  <a:ext cx=\"1905000\" cy=\"1905000\"/>\n\t\t\t\t</a:xfrm>\n\t\t\t\t<a:prstGeom prst=\"rect\">\n\t\t\t\t  <a:avLst/>\n\t\t\t\t</a:prstGeom>\n\t\t\t  </pic:spPr>\n\t\t\t</pic:pic>\n\t\t  </a:graphicData>\n\t\t  </a:graphic>\n\t\t  </wp:inline>\n\t\t  </w:drawing>\n\t</w:r>\n</w:p><w:sectPr><w:headerReference w:type=\"default\" r:id=\"hId0\"/><w:type w:val=\"continuous\"/><w:pgSz w:w=\"12240\" w:h=\"15840\" w:orient=\"portrait\"/><w:pgMar w:top=\"2810\" w:left=\"1800\" w:right=\"1800\" w:bottom=\"1440\"/><w:cols w:num=\"1\" w:sep=\"off\" w:equalWidth=\"1\"/></w:sectPr></w:body></w:document>""")
	# 	#
	# 	# expect(documentContent).to.contain('align')
	# 	# expect(documentContent).to.contain('center')

	# 	fs.writeFile("test_center.docx",zip.generate({type:"nodebuffer"}));


	# it 'should work with loops',()->
	# 	name='imageLoopExample.docx'

	# 	imageModule=new ImageModule({centered:true})
	# 	docX[name].attachModule(imageModule)

	# 	out=docX[name]
	# 		.load(docX[name].loadedContent)
	# 		.setData({images:['examples/image.png','examples/image2.png']})

	# 	out
	# 		.render()

	# 	zip=out.getZip()

	# 	imageFile=zip.files['word/media/image_generated_1.png']
	# 	expect(imageFile?).to.equal(true)
	# 	expect(imageFile.asText().length).to.equal(17417)

	# 	imageFile2=zip.files['word/media/image_generated_2.png']
	# 	expect(imageFile2?).to.equal(true)
	# 	expect(imageFile2.asText().length).to.equal(7177)

	# 	relsFile=zip.files['word/_rels/document.xml.rels']
	# 	expect(relsFile?).to.equal(true)
	# 	relsFileContent=relsFile.asText()
	# 	expect(relsFileContent).to.equal("""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\" Target=\"numbering.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\" Target=\"footnotes.xml\"/><Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes\" Target=\"endnotes.xml\"/><Relationship Id=\"hId0\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header0.xml\"/><Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image_generated_1.png\"/><Relationship Id=\"rId7\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image_generated_2.png\"/></Relationships>""")

	# 	documentFile=zip.files['word/document.xml']
	# 	expect(documentFile?).to.equal(true)
	# 	documentContent=documentFile.asText()
	# 	# expect(documentContent).to.equal("""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wx=\"http://schemas.microsoft.com/office/word/2003/auxHint\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:pPr></w:pPr><w:r><w:rPr></w:rPr><w:t xml:space=\"preserve\"></w:t></w:r></w:p><w:p><w:pPr></w:pPr></w:p><w:p><w:pPr></w:pPr><w:r><w:rPr></w:rPr><w:drawing>\n  <wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">\n    <wp:extent cx=\"1905000\" cy=\"1905000\"/>\n    <wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>\n    <wp:docPr id=\"2\" name=\"Image 2\" descr=\"description\"/>\n    <wp:cNvGraphicFramePr>\n      <a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/>\n    </wp:cNvGraphicFramePr>\n    <a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n      <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\n        <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\n          <pic:nvPicPr>\n            <pic:cNvPr id=\"0\" name=\"Picture 1\" descr=\"description\"/>\n            <pic:cNvPicPr>\n              <a:picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\"/>\n            </pic:cNvPicPr>\n          </pic:nvPicPr>\n          <pic:blipFill>\n            <a:blip r:embed=\"rId6\">\n              <a:extLst>\n                <a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\">\n                  <a14:useLocalDpi xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/>\n                </a:ext>\n              </a:extLst>\n            </a:blip>\n            <a:srcRect/>\n            <a:stretch>\n              <a:fillRect/>\n            </a:stretch>\n          </pic:blipFill>\n          <pic:spPr bwMode=\"auto\">\n            <a:xfrm>\n              <a:off x=\"0\" y=\"0\"/>\n              <a:ext cx=\"1905000\" cy=\"1905000\"/>\n            </a:xfrm>\n            <a:prstGeom prst=\"rect\">\n              <a:avLst/>\n            </a:prstGeom>\n            <a:noFill/>\n            <a:ln>\n              <a:noFill/>\n            </a:ln>\n          </pic:spPr>\n        </pic:pic>\n      </a:graphicData>\n    </a:graphic>\n  </wp:inline>\n</w:drawing></w:r></w:p><w:p><w:pPr></w:pPr></w:p><w:p><w:pPr></w:pPr><w:r><w:rPr></w:rPr><w:t xml:space=\"preserve\"></w:t></w:r></w:p><w:p><w:pPr></w:pPr></w:p><w:p><w:pPr></w:pPr><w:r><w:rPr></w:rPr><w:drawing>\n  <wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">\n    <wp:extent cx=\"1905000\" cy=\"1905000\"/>\n    <wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>\n    <wp:docPr id=\"2\" name=\"Image 2\" descr=\"description\"/>\n    <wp:cNvGraphicFramePr>\n      <a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/>\n    </wp:cNvGraphicFramePr>\n    <a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n      <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\n        <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\n          <pic:nvPicPr>\n            <pic:cNvPr id=\"0\" name=\"Picture 1\" descr=\"description\"/>\n            <pic:cNvPicPr>\n              <a:picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\"/>\n            </pic:cNvPicPr>\n          </pic:nvPicPr>\n          <pic:blipFill>\n            <a:blip r:embed=\"rId7\">\n              <a:extLst>\n                <a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\">\n                  <a14:useLocalDpi xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/>\n                </a:ext>\n              </a:extLst>\n            </a:blip>\n            <a:srcRect/>\n            <a:stretch>\n              <a:fillRect/>\n            </a:stretch>\n          </pic:blipFill>\n          <pic:spPr bwMode=\"auto\">\n            <a:xfrm>\n              <a:off x=\"0\" y=\"0\"/>\n              <a:ext cx=\"1905000\" cy=\"1905000\"/>\n            </a:xfrm>\n            <a:prstGeom prst=\"rect\">\n              <a:avLst/>\n            </a:prstGeom>\n            <a:noFill/>\n            <a:ln>\n              <a:noFill/>\n            </a:ln>\n          </pic:spPr>\n        </pic:pic>\n      </a:graphicData>\n    </a:graphic>\n  </wp:inline>\n</w:drawing></w:r></w:p><w:p><w:pPr></w:pPr></w:p><w:p><w:pPr></w:pPr><w:r><w:rPr></w:rPr><w:t xml:space=\"preserve\"></w:t></w:r><w:bookmarkStart w:id=\"20\" w:name=\"_GoBack\"/><w:bookmarkEnd w:id=\"20\"/></w:p><w:sectPr><w:headerReference w:type=\"default\" r:id=\"hId0\"/><w:type w:val=\"continuous\"/><w:pgSz w:w=\"12240\" w:h=\"15840\" w:orient=\"portrait\"/><w:pgMar w:top=\"2810\" w:left=\"1800\" w:right=\"1800\" w:bottom=\"1440\"/><w:cols w:num=\"1\" w:sep=\"off\" w:equalWidth=\"1\"/></w:sectPr></w:body></w:document>""")

	# 	buffer=zip.generate({type:"nodebuffer"})
	# 	fs.writeFile("test_multi.docx",buffer);
	
	# it 'should work with image in header/footer',()->
	# 	name='imageHeaderFooterExample.docx'
	# 	imageModule=new ImageModule({centered:false})
	# 	docX[name].attachModule(imageModule)
	# 	out=docX[name]
	# 		.load(docX[name].loadedContent)
	# 		.setData({image:'examples/image.png'})
	# 		.render()

	# 	zip=out.getZip()

	# 	imageFile=zip.files['word/media/image_generated_1.png']
	# 	expect(imageFile?).to.equal(true)
	# 	expect(imageFile.asText().length).to.equal(17417)

	# 	imageFile2=zip.files['word/media/image_generated_2.png']
	# 	expect(imageFile2?).to.equal(true)
	# 	expect(imageFile2.asText().length).to.equal(17417)

	# 	relsFile=zip.files['word/_rels/document.xml.rels']
	# 	expect(relsFile?).to.equal(true)
	# 	relsFileContent=relsFile.asText()
	# 	expect(relsFileContent).to.equal("""<?xml version="1.0" encoding="UTF-8"?>
	# 		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
	# 		</Relationships>""")

	# 	headerRelsFile=zip.files['word/_rels/header1.xml.rels']
	# 	expect(headerRelsFile?).to.equal(true)
	# 	headerRelsFileContent=headerRelsFile.asText()
	# 	expect(headerRelsFileContent).to.equal("""<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
	# 		<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image_generated_2.png"/></Relationships>""")

	# 	footerRelsFile=zip.files['word/_rels/footer1.xml.rels']
	# 	expect(footerRelsFile?).to.equal(true)
	# 	footerRelsFileContent=footerRelsFile.asText()
	# 	expect(footerRelsFileContent).to.equal("""<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
	# 		<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image_generated_1.png"/></Relationships>""")

	# 	documentFile=zip.files['word/document.xml']
	# 	expect(documentFile?).to.equal(true)
	# 	documentContent=documentFile.asText()
	# 	expect(documentContent).to.equal("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	# 		<w:document xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body><w:p><w:pPr><w:pStyle w:val="Normal"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:footerReference w:type="default" r:id="rId3"/><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1800" w:right="1800" w:header="720" w:top="2810" w:footer="1440" w:bottom="2003" w:gutter="0"/><w:pgNumType w:fmt="decimal"/><w:formProt w:val="false"/><w:textDirection w:val="lrTb"/><w:docGrid w:type="default" w:linePitch="249" w:charSpace="2047"/></w:sectPr></w:body></w:document>""")

	# 	fs.writeFile("test_header_footer.docx",zip.generate({type:"nodebuffer"}));


