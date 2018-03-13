DocUtils = require('./docUtils')
module.exports = class ChartMaker
	getTemplateTop: (chartType) ->
		return """
			<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<c:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
				<c:lang val="ru-RU"/>
				<c:chart>
					#{if @options.title then "" else "<c:autoTitleDeleted val=\"1\"/>"}
					<c:plotArea>
						<c:layout/>
						<c:barChart>
						<c:radarStyle val="marker"/>
						<c:varyColors val="1"/>
		"""

	getFormatCode: () ->
		if @options.axis.x.type == 'date'
			return "<c:formatCode>m/d/yyyy</c:formatCode>" 
		else 
			return ""
	getLineTemplate: (chartType, line, i) ->
		result = """
			<c:ser>
				<c:idx val="#{i}"/>
				<c:order val="#{i}"/>
				<c:tx>
					<c:v>#{line.name}</c:v>
				</c:tx>
		"""
		if (i==1)
			result += "<c:spPr><a:ln w=\"28800\"><a:solidFill><a:srgbClr val=\"990000\"/></a:solidFill><a:prstDash val=\"dash\"/><a:round/></a:ln></c:spPr>\n"

		result += """
				<c:marker>
					<c:symbol val="none"/>
				</c:marker>
				<c:cat>
					<c:#{@ref}>
						<c:#{@cache}>
							#{@getFormatCode()}
							<c:ptCount val="#{line.data.length}"/>

		"""
		for elem, i in line.data
			result += """
				<c:pt idx="#{i}">
					<c:v>#{elem.x}</c:v>
				</c:pt>
			"""
		result += """
						</c:#{@cache}>
					</c:#{@ref}>
				</c:cat>
				<c:val>
					<c:numRef>
						<c:numCache>
							<c:formatCode>General</c:formatCode>
							<c:ptCount val="#{line.data.length}"/>
		"""
		for elem, i in line.data
			result += """
				<c:pt idx="#{i}">
					<c:v>#{elem.y}</c:v>
				</c:pt>
			"""
		result += """
						</c:numCache>
					</c:numRef>
				</c:val>
			</c:ser>
		"""
		return result
	id1: 142309248,
	id2: 142310784
	getScaling: (opts) ->
		"""
		<c:scaling>
			<c:orientation val="#{opts.orientation}"/>
			#{if opts.max != undefined then "<c:max val=\"#{opts.max}\"/>" else ""}
			#{if opts.min != undefined then "<c:min val=\"#{opts.min}\"/>" else ""}
		</c:scaling>
		"""
	getAxOpts: () ->
		return """
		<c:axId val="#{@id1}"/>
		#{@getScaling(@options.axis.x)}
		<c:axPos val="b"/>

		<c:delete val="0"/>
    <c:majorGridlines/>
                <c:majorTickMark val="cross"/>
                <c:minorTickMark val="cross"/>
		<c:tickLblPos val="nextTo"/>
		<c:txPr>
			<a:bodyPr/>
			<a:lstStyle/>
			<a:p>
				<a:pPr>
					<a:defRPr sz="800"/>
				</a:pPr>
				<a:endParaRPr lang="ru-RU"/>
			</a:p>
		</c:txPr>
		<c:crossAx val="#{@id2}"/>
		<c:crosses val="autoZero"/>
		<c:auto val="1"/>
		<c:lblOffset val="100"/>
    <c:noMultiLvlLbl val="1"/>
		"""
	getCatAx: () ->
		return """
		<c:catAx>
			#{@getAxOpts()}
			<c:lblAlgn val="ctr"/>
		</c:catAx>
		"""
	getDateAx: () ->
		return """
		<c:dateAx>
			#{@getAxOpts()}
			<c:delete val="0"/>
			<c:numFmt formatCode="#{@options.axis.x.date.code}" sourceLinked="0"/>
			<c:majorTickMark val="out"/>
			<c:minorTickMark val="none"/>
			<c:baseTimeUnit val="days"/>
			<c:majorUnit val="#{@options.axis.x.date.step}"/>
			<c:majorTimeUnit val="#{@options.axis.x.date.unit}"/>
		</c:dateAx>
		"""
	getBorder: () ->
		unless @options.border
			return """
				<c:spPr>
					<a:ln>
						<a:noFill/>
					</a:ln>
				</c:spPr>
			"""
		else
			return ''
	getTemplateBottom: (chartType) ->
		result = """
							<c:marker val="1"/>
                <c:dLbls>
                    <c:showLegendKey val="0"/>
                    <c:showVal val="0"/>
                    <c:showCatName val="0"/>
                    <c:showSerName val="0"/>
                    <c:showPercent val="0"/>
                    <c:showBubbleSize val="0"/>
                </c:dLbls>
							<c:axId val="#{@id1}"/>
							<c:axId val="#{@id2}"/>
						</c:barChart>
		"""
		switch @options.axis.x.type
			when 'date'
				result += @getDateAx()
			else
				result += @getCatAx()
		result += """
						<c:valAx>
							<c:axId val="#{@id2}"/>
							#{@getScaling(@options.axis.y)}
							<c:axPos val="l"/>
							<c:delete val="0"/>
							#{if @options.grid then "<c:majorGridlines>
								<c:spPr>
								<a:ln>
								<a:solidFill>
								<a:schemeClr val=\"bg1\">
								<a:lumMod val=\"85000\"/>
								</a:schemeClr>
								\</a:solidFill>
								\</a:ln>
								\</c:spPr>
								\</c:majorGridlines>" else ""}
			<c:numFmt formatCode="General" sourceLinked="1"/>                <c:majorTickMark val="none"/>
			<c:minorTickMark val="none"/>
			<c:tickLblPos val="none"/>
			<c:spPr>
			<a:ln>
			<a:solidFill>
			<a:schemeClr val="bg1">
			<a:lumMod val="75000"/>
			</a:schemeClr>
			</a:solidFill>
			</a:ln>
			</c:spPr>
							<c:txPr>
								<a:bodyPr/>
								<a:lstStyle/>
								<a:p>
									<a:pPr>
										<a:defRPr sz="600"/>
									</a:pPr>
									<a:endParaRPr lang="ru-RU"/>
								</a:p>
							</c:txPr>
							<c:crossAx val="#{@id1}"/>
							<c:crosses val="autoZero"/>
							<c:crossBetween val="between"/>                <c:majorUnit val="1"/>
						</c:valAx>
					</c:plotArea>
					<c:legend>
						<c:legendPos val="#{@options.legend.position}"/>
						<c:layout/>
					</c:legend>
					<c:plotVisOnly val="1"/>
				</c:chart>
				#{@getBorder()}
			</c:chartSpace>
		"""
		return result
	constructor: (@zip, @options) ->
		if (@options.axis.x.type == 'date')
			@ref = "numRef"
			@cache = "numCache"
		else
			@ref = "strRef"
			@cache = "strCache"
			

	makeChartFile: (chartType, lines) ->
		result = @getTemplateTop(chartType)
		for line, i in lines
			result += @getLineTemplate(chartType, line, i)
		result += @getTemplateBottom(chartType)
		@chartContent = result
		return @chartContent

	writeFile: (path) ->
		@zip.file("word/charts/#{path}.xml", @chartContent, {})
		return
