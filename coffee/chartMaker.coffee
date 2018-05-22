DocUtils = require('./docUtils')
module.exports = class ChartMaker
	getTemplateTop: (chartType, title1, title2) ->
		switch chartType
			#**********TEMPLATE TOP RADAR*****************************
			when 'radar'
				return """
					<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
					<c:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
						<c:lang val="ru-RU"/>
						<c:chart>
							#{if @options.title then "" else "<c:autoTitleDeleted val=\"1\"/>"}
							<c:plotArea>
								<c:layout/>
								<c:radarChart>
								<c:radarStyle val="marker"/>
								<c:varyColors val="1"/>
				"""
			#*******************************************************

			#**********TEMPLATE TOP BAR*****************************
			when 'bar'
				return """
					<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
					<c:chartSpace
						xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
						xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
						xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
						xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart">
						<c:date1904 val="0"/>
						<c:lang val="en-US"/>
						<c:roundedCorners val="0"/>
						<mc:AlternateContent
							xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
							<mc:Choice Requires="c14"
								xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">
								<c14:style val="102"/>
							</mc:Choice>
							<mc:Fallback>
								<c:style val="2"/>
							</mc:Fallback>
						</mc:AlternateContent>
						<c:chart>
							<c:title>
								<c:tx>
									<c:rich>
										<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
										<a:lstStyle/>
										<a:p>
											<a:pPr>
												<a:defRPr sz="1000" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0">
													<a:solidFill>
														<a:schemeClr val="tx1">
															<a:lumMod val="65000"/>
															<a:lumOff val="35000"/>
														</a:schemeClr>
													</a:solidFill>
													<a:latin typeface="+mn-lt"/>
													<a:ea typeface="+mn-ea"/>
													<a:cs typeface="+mn-cs"/>
												</a:defRPr>
											</a:pPr>
											<a:r>
												<a:rPr lang="en-US" sz="1000"/>
												<a:t>#{title1}</a:t>
											</a:r>
											<a:r>
												<a:rPr lang="en-US" sz="1000" baseline="0"/>
												<a:t>#{title2}</a:t>
											</a:r>
											<a:endParaRPr lang="en-US" sz="1000"/>
										</a:p>
									</c:rich>
								</c:tx>
								<c:overlay val="0"/>
								<c:spPr>
									<a:noFill/>
									<a:ln>
										<a:noFill/>
									</a:ln>
									<a:effectLst/>
								</c:spPr>
								<c:txPr>
									<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
									<a:lstStyle/>
									<a:p>
										<a:pPr>
											<a:defRPr sz="1000" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0">
												<a:solidFill>
													<a:schemeClr val="tx1">
														<a:lumMod val="65000"/>
														<a:lumOff val="35000"/>
													</a:schemeClr>
												</a:solidFill>
												<a:latin typeface="+mn-lt"/>
												<a:ea typeface="+mn-ea"/>
												<a:cs typeface="+mn-cs"/>
											</a:defRPr>
										</a:pPr>
										<a:endParaRPr lang="en-US"/>
									</a:p>
								</c:txPr>
							</c:title>
							<c:autoTitleDeleted val="0"/>
							<c:plotArea>
								<c:layout/>
								<c:barChart>
									<c:barDir val="col"/>
									<c:grouping val="clustered"/>
									<c:varyColors val="0"/>
				"""
			#*******************************************************

#**********TEMPLATE TOP LINE*****************************
			when 'line'
				return """
					<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
					<c:chartSpace
						xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
						xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
						xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
						xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart">
						<c:date1904 val="0"/>
						<c:lang val="en-US"/>
						<c:roundedCorners val="0"/>
						<mc:AlternateContent
							xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
							<mc:Choice Requires="c14"
								xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">
								<c14:style val="105"/>
							</mc:Choice>
							<mc:Fallback>
								<c:style val="5"/>
							</mc:Fallback>
						</mc:AlternateContent>
						<c:chart>
							<c:title>
								<c:tx>
									<c:rich>
										<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
										<a:lstStyle/>
										<a:p>
											<a:pPr>
												<a:defRPr sz="1400" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0">
													<a:solidFill>
														<a:schemeClr val="tx1"/>
													</a:solidFill>
													<a:latin typeface="+mn-lt"/>
													<a:ea typeface="+mn-ea"/>
													<a:cs typeface="+mn-cs"/>
												</a:defRPr>
											</a:pPr>
											<a:r>
												<a:rPr lang="en-IE"/>
												<a:t>#{title1}</a:t>
											</a:r>
										</a:p>
									</c:rich>
								</c:tx>
								<c:overlay val="0"/>
								<c:spPr>
									<a:noFill/>
									<a:ln>
										<a:noFill/>
									</a:ln>
									<a:effectLst/>
								</c:spPr>
								<c:txPr>
									<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
									<a:lstStyle/>
									<a:p>
										<a:pPr>
											<a:defRPr sz="1400" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0">
												<a:solidFill>
													<a:schemeClr val="tx1"/>
												</a:solidFill>
												<a:latin typeface="+mn-lt"/>
												<a:ea typeface="+mn-ea"/>
												<a:cs typeface="+mn-cs"/>
											</a:defRPr>
										</a:pPr>
										<a:endParaRPr lang="en-US"/>
									</a:p>
								</c:txPr>
							</c:title>
							<c:autoTitleDeleted val="0"/>
							<c:plotArea>
								<c:layout/>
								<c:lineChart>
									<c:grouping val="standard"/>
									<c:varyColors val="0"/>
				"""
#*******************************************************


	getFormatCode: () ->
		if @options.axis.x.type == 'date'
			return "<c:formatCode>m/d/yyyy</c:formatCode>" 
		else 
			return ""


	getLineTemplate: (chartType, line, lineCounter) ->
		switch chartType
			#**********LINE TEMPLATE RADAR*****************************
			when 'radar'
				result = """
					<c:ser>
						<c:idx val="#{lineCounter}"/>
						<c:order val="#{lineCounter}"/>
						<c:tx>
							<c:v>#{line.name}</c:v>
						</c:tx>
				"""
				if (lineCounter==1)
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
			#*******************************************************

			#**********LINE TEMPLATE BAR*****************************
			when 'bar'
				result = """
					<c:ser>
						<c:idx val="#{lineCounter}"/>
						<c:order val="#{lineCounter}"/>
						<c:tx>
							<c:v>#{line.name}</c:v>
						</c:tx>
				"""
#				if (lineCounter==1)
#					result += "<c:spPr><a:ln w=\"28800\"><a:solidFill><a:srgbClr val=\"990000\"/></a:solidFill><a:prstDash val=\"dash\"/><a:round/></a:ln></c:spPr>\n"

				result += """
					<c:spPr>
						<a:solidFill>
							<a:schemeClr val="accent#{lineCounter + 1}"/>
						</a:solidFill>
						<a:ln>
							<a:noFill/>
						</a:ln>
						<a:effectLst/>
					</c:spPr>
					<c:invertIfNegative val="0"/>
					<c:dLbls>
						<c:spPr>
							<a:noFill/>
							<a:ln>
								<a:noFill/>
							</a:ln>
							<a:effectLst/>
						</c:spPr>
						<c:txPr>
							<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" lIns="38100" tIns="19050" rIns="38100" bIns="19050" anchor="ctr" anchorCtr="1">
								<a:spAutoFit/>
							</a:bodyPr>
							<a:lstStyle/>
							<a:p>
								<a:pPr>
									<a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
										<a:solidFill>
											<a:schemeClr val="tx1">
												<a:lumMod val="75000"/>
												<a:lumOff val="25000"/>
											</a:schemeClr>
										</a:solidFill>
										<a:latin typeface="+mn-lt"/>
										<a:ea typeface="+mn-ea"/>
										<a:cs typeface="+mn-cs"/>
									</a:defRPr>
								</a:pPr>
								<a:endParaRPr lang="en-US"/>
							</a:p>
						</c:txPr>
						<c:dLblPos val="outEnd"/>
						<c:showLegendKey val="0"/>
						<c:showVal val="1"/>
						<c:showCatName val="0"/>
						<c:showSerName val="0"/>
						<c:showPercent val="0"/>
						<c:showBubbleSize val="0"/>
						<c:showLeaderLines val="0"/>
						<c:extLst>
							<c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}"
								xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">
								<c15:showLeaderLines val="1"/>
								<c15:leaderLines>
									<c:spPr>
										<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
											<a:solidFill>
												<a:schemeClr val="tx1">
													<a:lumMod val="35000"/>
													<a:lumOff val="65000"/>
												</a:schemeClr>
											</a:solidFill>
											<a:round/>
										</a:ln>
										<a:effectLst/>
									</c:spPr>
								</c15:leaderLines>
							</c:ext>
						</c:extLst>
					</c:dLbls>
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
						<c:extLst>
							<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"
								xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">
								<c16:uniqueId val="{00000000-9819-4A27-9D24-58E0D8F9109B}"/>
							</c:ext>
						</c:extLst>
					</c:ser>
				"""
		#*******************************************************

#**********LINE TEMPLATE LINE*****************************
			when 'line'
				result = """
					<c:ser>
						<c:idx val="#{lineCounter}"/>
						<c:order val="#{lineCounter}"/>
					<c:tx>
						<c:strRef>
							<c:strCache>
								<c:ptCount val="1"/>
								<c:pt idx="0">
									<c:v>Series 1</c:v>
								</c:pt>
							</c:strCache>
						</c:strRef>
					</c:tx>
				"""
				#				if (lineCounter==1)
				#					result += "<c:spPr><a:ln w=\"28800\"><a:solidFill><a:srgbClr val=\"990000\"/></a:solidFill><a:prstDash val=\"dash\"/><a:round/></a:ln></c:spPr>\n"

				result += """
					<c:spPr>
						<a:ln w="28575" cap="rnd">
							<a:solidFill>
								<a:schemeClr val="accent1"/>
							</a:solidFill>
							<a:round/>
						</a:ln>
						<a:effectLst/>
					</c:spPr>
					<c:marker>
						<c:symbol val="none"/>
					</c:marker>
					<c:cat>
						<c:numRef>
							<c:numCache>
									<c:formatCode>General</c:formatCode>
								<c:ptCount val="#{line.data.length}"/>

				"""
				for elem, i in line.data
					result += """
						<c:pt idx="#{i}">
							<c:v>#{elem.x}</c:v>
						</c:pt>
					"""
				result += """
								</c:numCache>
							</c:numRef>
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
						<c:extLst>
							<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"
								xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">
								<c16:uniqueId val="{00000000-9819-4A27-9D24-58E0D8F9109B}"/>
							</c:ext>
						</c:extLst>
					</c:ser>
				"""
		#*******************************************************

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


	getAxOpts: (chartType) ->
		switch chartType
			#**********getAxOpts RADAR*****************************
			when 'radar'
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
			#*******************************************************

			#**********getAxOpts BAR*****************************
			when 'bar'
				return """
				<c:axId val="#{@id1}"/>
				#{@getScaling(@options.axis.x)}
				<c:axPos val="b"/>

				<c:delete val="0"/>
				<c:majorTickMark val="none"/>
				<c:minorTickMark val="none"/>
				<c:tickLblPos val="none"/>
				<c:spPr>
					<a:noFill/>
					<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
						<a:solidFill>
							<a:schemeClr val="tx1">
								<a:lumMod val="15000"/>
								<a:lumOff val="85000"/>
							</a:schemeClr>
						</a:solidFill>
						<a:round/>
					</a:ln>
					<a:effectLst/>
				</c:spPr>
				<c:txPr>
					<a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
					<a:lstStyle/>
					<a:p>
						<a:pPr>
							<a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
								<a:solidFill>
									<a:schemeClr val="tx1">
										<a:lumMod val="65000"/>
										<a:lumOff val="35000"/>
									</a:schemeClr>
								</a:solidFill>
								<a:latin typeface="+mn-lt"/>
								<a:ea typeface="+mn-ea"/>
								<a:cs typeface="+mn-cs"/>
							</a:defRPr>
						</a:pPr>
						<a:endParaRPr lang="en-US"/>
					</a:p>
				</c:txPr>
				<c:crossAx val="#{@id2}"/>
				<c:crosses val="autoZero"/>
				<c:auto val="1"/>
				<c:lblOffset val="100"/>
				<c:noMultiLvlLbl val="1"/>
			"""
			#*******************************************************

			#**********getAxOpts LINE*****************************
			when 'line'
							return """
							<c:axId val="#{@id1}"/>
							#{@getScaling(@options.axis.x)}
							<c:axPos val="b"/>

						<c:delete val="0"/>
										<c:majorTickMark val="out"/>
										<c:minorTickMark val="none"/>
										<c:tickLblPos val="nextTo"/>
										<c:spPr>
											<a:noFill/>
											<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
												<a:solidFill>
													<a:schemeClr val="tx1">
														<a:lumMod val="15000"/>
														<a:lumOff val="85000"/>
													</a:schemeClr>
												</a:solidFill>
												<a:round/>
											</a:ln>
											<a:effectLst/>
										</c:spPr>
										<c:txPr>
											<a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
											<a:lstStyle/>
											<a:p>
												<a:pPr>
													<a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
														<a:solidFill>
															<a:schemeClr val="tx1"/>
														</a:solidFill>
														<a:latin typeface="+mn-lt"/>
														<a:ea typeface="+mn-ea"/>
														<a:cs typeface="+mn-cs"/>
													</a:defRPr>
												</a:pPr>
												<a:endParaRPr lang="en-US"/>
											</a:p>
										</c:txPr>
							<c:crossAx val="#{@id2}"/>
							<c:crosses val="autoZero"/>
							<c:auto val="1"/>
							<c:lblOffset val="100"/>
							<c:noMultiLvlLbl val="0"/>
			"""
			#*******************************************************


	getCatAx: (chartType) ->
		switch chartType
			#**********getCatAx RADAR*****************************
			when 'radar'
				return """
					<c:catAx>
						#{@getAxOpts(chartType)}
						<c:lblAlgn val="ctr"/>
					</c:catAx>
					"""
			#*******************************************************

			#**********getCatAx BAR*****************************
			when 'bar'
				return """
					<c:catAx>
						#{@getAxOpts(chartType)}
						<c:lblAlgn val="ctr"/>
					</c:catAx>
					"""
			#*******************************************************

			#**********getCatAx LINE*****************************
			when 'line'
							return """
								<c:catAx>
									#{@getAxOpts(chartType)}
									<c:lblAlgn val="ctr"/>
								</c:catAx>
								"""
			#*******************************************************


	getDateAx: (chartType) ->
		return """
		<c:dateAx>
			#{@getAxOpts(chartType)}
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
		switch chartType
			#**********TEMPLATE BOTTOM RADAR*****************************
			when 'radar'
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
								</c:radarChart>
				"""
				switch @options.axis.x.type
					when 'date'
						result += @getDateAx(chartType)
					else
						result += @getCatAx(chartType)
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
			#*******************************************************

			#**********TEMPLATE BOTTOM BAR*****************************
			when 'bar'
				result = """
									<c:dLbls>
										<c:dLblPos val="outEnd"/>
										<c:showLegendKey val="0"/>
										<c:showVal val="1"/>
										<c:showCatName val="0"/>
										<c:showSerName val="0"/>
										<c:showPercent val="0"/>
										<c:showBubbleSize val="0"/>
									</c:dLbls>
									<c:gapWidth val="219"/>
									<c:overlap val="-27"/>
									<c:axId val="#{@id1}"/>
									<c:axId val="#{@id2}"/>
								</c:barChart>
				"""
				switch @options.axis.x.type
					when 'date'
						result += @getDateAx(chartType)
					else
						result += @getCatAx(chartType)
				result += """
												<c:valAx>
													<c:axId val="#{@id2}"/>
													#{@getScaling(@options.axis.y)}
													<c:axPos val="l"/>
													<c:delete val="0"/>
													<c:majorGridlines>
														<c:spPr>
															<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
																<a:noFill/>
																<a:round/>
															</a:ln>
															<a:effectLst/>
														</c:spPr>
													</c:majorGridlines>
													<c:numFmt formatCode="General" sourceLinked="1"/>
													<c:majorTickMark val="none"/>
													<c:minorTickMark val="none"/>
													<c:tickLblPos val="nextTo"/>
													<c:spPr>
														<a:noFill/>
														<a:ln>
															<a:noFill/>
														</a:ln>
														<a:effectLst/>
													</c:spPr>
													<c:txPr>
														<a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
														<a:lstStyle/>
														<a:p>
															<a:pPr>
																<a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
																	<a:solidFill>
																		<a:schemeClr val="tx1">
																			<a:lumMod val="65000"/>
																			<a:lumOff val="35000"/>
																		</a:schemeClr>
																	</a:solidFill>
																	<a:latin typeface="+mn-lt"/>
																	<a:ea typeface="+mn-ea"/>
																	<a:cs typeface="+mn-cs"/>
																</a:defRPr>
															</a:pPr>
															<a:endParaRPr lang="en-US"/>
														</a:p>
													</c:txPr>
													<c:crossAx val="#{@id1}"/>
													<c:crosses val="autoZero"/>
													<c:crossBetween val="between"/>
												</c:valAx>
												<c:spPr>
													<a:noFill/>
													<a:ln>
														<a:noFill/>
													</a:ln>
													<a:effectLst/>
												</c:spPr>
											</c:plotArea>
											<c:legend>
												<c:overlay val="0"/>
												<c:spPr>
													<a:noFill/>
													<a:ln>
														<a:noFill/>
													</a:ln>
													<a:effectLst/>
												</c:spPr>
												<c:txPr>
													<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
													<a:lstStyle/>
													<a:p>
														<a:pPr>
															<a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
																<a:solidFill>
																	<a:schemeClr val="tx1">
																		<a:lumMod val="65000"/>
																		<a:lumOff val="35000"/>
																	</a:schemeClr>
																</a:solidFill>
																<a:latin typeface="+mn-lt"/>
																<a:ea typeface="+mn-ea"/>
																<a:cs typeface="+mn-cs"/>
															</a:defRPr>
														</a:pPr>
														<a:endParaRPr lang="en-US"/>
													</a:p>
												</c:txPr>
											<c:legendPos val="#{@options.legend.position}"/>
										</c:legend>
										<c:plotVisOnly val="1"/>
										<c:dispBlanksAs val="gap"/>
										<c:extLst>
											<c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}"
												xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart">
												<c16r3:dataDisplayOptions16>
													<c16r3:dispNaAsBlank val="1"/>
												</c16r3:dataDisplayOptions16>
											</c:ext>
										</c:extLst>
										<c:showDLblsOverMax val="0"/>
									</c:chart>
									<c:spPr>
										<a:solidFill>
											<a:schemeClr val="bg1"/>
										</a:solidFill>
										<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
											<a:noFill/>
											<a:round/>
										</a:ln>
										<a:effectLst/>
									</c:spPr>
									<c:txPr>
										<a:bodyPr/>
										<a:lstStyle/>
										<a:p>
											<a:pPr>
												<a:defRPr/>
											</a:pPr>
											<a:endParaRPr lang="en-US"/>
										</a:p>
									</c:txPr>
								</c:chartSpace>
			"""
		#*******************************************************

		#**********TEMPLATE BOTTOM LINE*****************************
			when 'line'
						result = """
											<c:dLbls>
												<c:showLegendKey val="0"/>
												<c:showVal val="0"/>
												<c:showCatName val="0"/>
												<c:showSerName val="0"/>
												<c:showPercent val="0"/>
												<c:showBubbleSize val="0"/>
											</c:dLbls>
											<c:smooth val="0"/>
											<c:axId val="#{@id1}"/>
											<c:axId val="#{@id2}"/>
										</c:lineChart>
						"""
						switch @options.axis.x.type
							when 'date'
								result += @getDateAx(chartType)
							else
								result += @getCatAx(chartType)
						result += """
														<c:valAx>
															<c:axId val="#{@id2}"/>
															<c:scaling>
																<c:orientation val="minMax"/>
															</c:scaling>
															<c:axPos val="l"/>
															<c:delete val="0"/>
															<c:majorGridlines>
																<c:spPr>
																	<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
																		<a:noFill/>
																		<a:round/>
																	</a:ln>
																	<a:effectLst/>
																</c:spPr>
															</c:majorGridlines>
															<c:numFmt formatCode="General" sourceLinked="1"/>
															<c:majorTickMark val="none"/>
															<c:minorTickMark val="none"/>
															<c:tickLblPos val="nextTo"/>
															<c:spPr>
																<a:noFill/>
																<a:ln>
																	<a:solidFill>
																		<a:schemeClr val="bg1">
																			<a:lumMod val="75000"/>
																		</a:schemeClr>
																	</a:solidFill>
																</a:ln>
																<a:effectLst/>
															</c:spPr>
															<c:txPr>
																<a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
																<a:lstStyle/>
																<a:p>
																	<a:pPr>
																		<a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
																			<a:ln>
																				<a:noFill/>
																			</a:ln>
																			<a:solidFill>
																				<a:schemeClr val="tx1"/>
																			</a:solidFill>
																			<a:latin typeface="+mn-lt"/>
																			<a:ea typeface="+mn-ea"/>
																			<a:cs typeface="+mn-cs"/>
																		</a:defRPr>
																	</a:pPr>
																	<a:endParaRPr lang="en-US"/>
																</a:p>
															</c:txPr>
															<c:crossAx val="#{@id1}"/>
															<c:crosses val="autoZero"/>
															<c:crossBetween val="midCat"/>
														</c:valAx>
														<c:spPr>
															<a:noFill/>
															<a:ln>
																<a:noFill/>
															</a:ln>
															<a:effectLst/>
														</c:spPr>
													</c:plotArea>
												<c:plotVisOnly val="1"/>
												<c:dispBlanksAs val="gap"/>
												<c:extLst>
													<c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}"
														xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart">
														<c16r3:dataDisplayOptions16>
															<c16r3:dispNaAsBlank val="1"/>
														</c16r3:dataDisplayOptions16>
													</c:ext>
												</c:extLst>
												<c:showDLblsOverMax val="0"/>
											</c:chart>
											<c:spPr>
												<a:solidFill>
													<a:schemeClr val="bg1"/>
												</a:solidFill>
												<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
													<a:noFill/>
													<a:round/>
												</a:ln>
												<a:effectLst/>
											</c:spPr>
											<c:txPr>
												<a:bodyPr/>
												<a:lstStyle/>
												<a:p>
													<a:pPr>
														<a:defRPr>
															<a:solidFill>
																<a:schemeClr val="tx1"/>
															</a:solidFill>
														</a:defRPr>
													</a:pPr>
													<a:endParaRPr lang="en-US"/>
												</a:p>
											</c:txPr>
										</c:chartSpace>
			"""
			#*******************************************************
		return result


	constructor: (@zip, @options) ->
		if (@options.axis.x.type == 'date')
			@ref = "numRef"
			@cache = "numCache"
		else
			@ref = "strRef"
			@cache = "strCache"
			

	makeChartFile: (chart) ->
		result = @getTemplateTop(chart.chartType, chart.title1, chart.title2)
		for line, i in chart.lines
			result += @getLineTemplate(chart.chartType, line, i)
		result += @getTemplateBottom(chart.chartType)
		@chartContent = result
		return @chartContent

	writeFile: (path) ->
		@zip.file("word/charts/#{path}.xml", @chartContent, {})
		return
