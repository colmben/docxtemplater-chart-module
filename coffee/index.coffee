SubContent = require('docxtemplater').SubContent
ChartManager = require('./chartManager')
ChartMaker = require('./chartMaker')

class ChartModule
	###*
	 * self name for self-identification, variable for fast changing;
	 * @type {String}
	###
	name: 'chart'

	###*
	 * initialize options with empty object if not recived
	 * @manager = ModuleManager instance
	 * @param  {Object} @options params for the module
	###
	constructor: (@options = {}) ->

	handleEvent: (event, eventData) ->
		if (event == 'rendering-file')
			@renderingFileName = eventData;
			# console.log(renderingFileName)
			gen = @manager.getInstance('gen');
			@chartManager = new ChartManager(gen.zip, @renderingFileName)
			@chartManager.loadChartRels();

		else if (event == 'rendered')
			@finished()

	get: (data) ->
		# console.log('get data: ' + data);
		if data == 'loopType'
			templaterState = @manager.getInstance('templaterState')
			# console.log(templaterState.textInsideTag)
			if templaterState.textInsideTag[0] == '$'
				return @name
		return null

	handle: (type, data) ->
		if (type == 'replaceTag' and data == @name)
			# console.log('handle')
			@replaceTag()
		return null

	finished: () ->

	on: (event, data) ->
		if event == 'error'
			throw data

	replaceBy: (text, outsideElement) ->
		console.log('In replaceBy, text, outsideElement',text, outsideElement)
		xmlTemplater = @manager.getInstance('xmlTemplater')
		templaterState = @manager.getInstance('templaterState')
		console.log('xmlTemplater.content :', (xmlTemplater.content))
		console.log('Sub content :', new SubContent(xmlTemplater.content))
		subContent = new SubContent(xmlTemplater.content)
			.getInnerTag(templaterState)
			.getOuterXml(outsideElement)
		console.log('In replaceBy, subContent',subContent)
		xmlTemplater.replaceXml(subContent,text)

	convertPixelsToEmus: (pixel) ->
		Math.round(pixel * 9525)

	extendDefaults: (options) ->
		deepMerge = (target, source) ->
			for key of source
				original = target[key]
				next = source[key]
				if original and next and typeof next == 'object'
					deepMerge original, next
				else
					target[key] = next
			return target
#			width and height defaults are used for the radar diagrams
		defaultOptions = {
			width: 6477000 / 9525,
			height: 4152900 / 9525,
			grid: true,
			border: false,
			title: false, # works only for single-line charts
			format: 'General', # used in bar chart only currently
			legend: {
				position: 'r', # 'l', 'r', 'b', 't'
			},
			axis: {
				x: {
					orientation: 'minMax', # 'maxMin'
					min: undefined, # number
					max: undefined,
					type: undefined, # 'date'
					date: {
						format: 'unix',
						code: 'mm/yy', # "m/yy;@"
						unit: 'months', # "days"
						step: '1'
					}
				},
				y: {
					orientation: 'minMax',
					min: undefined,
					max: undefined
				}
			}
		}
		result = deepMerge({}, defaultOptions);
		result = deepMerge(result, options);
		return result;


	convertUnixTo1900: (chartData, axName) ->
		unixTo1900 = (value) ->
			return Math.round(value / 86400 + 25569)
		convertOption = (name) ->
			if (chartData.options.axis[axName][name])
				chartData.options.axis[axName][name] = unixTo1900(chartData.options.axis[axName][name])
		convertOption('min');
		convertOption('max');
		for line in chartData.lines
			for data in line.data
				data[axName] = unixTo1900(data[axName])
		return chartData

	replaceTag: () ->
		scopeManager = @manager.getInstance('scopeManager')
		templaterState = @manager.getInstance('templaterState')
		gen = @manager.getInstance('gen');

		tag = templaterState.textInsideTag.substr(1) # tag to be replaced
		console.log("Chart tag : ", tag);
		chartData = scopeManager.getValueFromScope(tag) # data to build chart from
		console.log("Chart chartData : ", chartData)

		# exit gracefully if no chartData, required when outer loop data doesn't exist
		return if !chartData?

		# create a unique filename so we can have multiple charts from one tag, via the loop functionality in docxtemplater
		# Note the +1 isn't really required, it just makes the number the same as the associated rel, handy for debugging the resulting docx
		filename = tag + (this.chartManager.maxRid + 1);


		imageRels = @chartManager.loadChartRels()

		return unless imageRels # break if no Relationships loaded
		chartId = @chartManager.addChartRels(filename)

		options = @extendDefaults(chartData.options)

		for name of options.axis
			ax = options.axis[name]
			if ax.type == 'date' and ax[ax.type].format == 'unix'
				chartData = @convertUnixTo1900(chartData, name)

		chart = new ChartMaker(gen.zip, options)
		chart.makeChartFile(chartData)
		chart.writeFile(filename)


		tagXml = @manager.getInstance('xmlTemplater').fileTypeConfig.tagsXmlArray[0]
		console.log("Chart tagXmlArray : ", @manager.getInstance('xmlTemplater').fileTypeConfig.tagsXmlArray)

		newText = @getChartXml({
			chartID: chartId,
			width: @convertPixelsToEmus(options.width),
			height: @convertPixelsToEmus(options.height)
		})
		console.log('Chart newText, tagXml', newText, tagXml);
		@replaceBy(newText, tagXml)

	getChartXml: ({chartID, width, height}) ->
		return """
			<w:drawing>
				<wp:inline distT="0" distB="0" distL="0" distR="0">
					<wp:extent cx="#{width}" cy="#{height}"/>
					<wp:effectExtent b="0" l="0" r="0" t="0"/>
					<wp:docPr id="1" name="Диаграмма 1"/>
					<wp:cNvGraphicFramePr/>
					<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
						<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
							<c:chart r:id="rId#{chartID}" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
						</a:graphicData>
					</a:graphic>
				</wp:inline>
			</w:drawing>
		"""



module.exports = ChartModule
