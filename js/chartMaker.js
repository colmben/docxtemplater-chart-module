var ChartMaker, DocUtils;

DocUtils = require('./docUtils');

module.exports = ChartMaker = (function() {
  ChartMaker.prototype.getTemplateTop = function(chartType, title1, title2) {
    switch (chartType) {
      case 'radar':
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<c:chartSpace xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n	<c:lang val=\"ru-RU\"/>\n	<c:chart>\n		" + (this.options.title ? "" : "<c:autoTitleDeleted val=\"1\"/>") + "\n		<c:plotArea>\n			<c:layout/>\n			<c:radarChart>\n			<c:radarStyle val=\"marker\"/>\n			<c:varyColors val=\"1\"/>";
      case 'bar':
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<c:chartSpace\n	xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"\n	xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"\n	xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"\n	xmlns:c16r2=\"http://schemas.microsoft.com/office/drawing/2015/06/chart\">\n	<c:date1904 val=\"0\"/>\n	<c:lang val=\"en-US\"/>\n	<c:roundedCorners val=\"0\"/>\n	<mc:AlternateContent\n		xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">\n		<mc:Choice Requires=\"c14\"\n			xmlns:c14=\"http://schemas.microsoft.com/office/drawing/2007/8/2/chart\">\n			<c14:style val=\"102\"/>\n		</mc:Choice>\n		<mc:Fallback>\n			<c:style val=\"2\"/>\n		</mc:Fallback>\n	</mc:AlternateContent>\n	<c:chart>\n		<c:title>\n			<c:tx>\n				<c:rich>\n					<a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n					<a:lstStyle/>\n					<a:p>\n						<a:pPr>\n							<a:defRPr sz=\"1000\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" spc=\"0\" baseline=\"0\">\n								<a:solidFill>\n									<a:schemeClr val=\"tx1\">\n										<a:lumMod val=\"65000\"/>\n										<a:lumOff val=\"35000\"/>\n									</a:schemeClr>\n								</a:solidFill>\n								<a:latin typeface=\"+mn-lt\"/>\n								<a:ea typeface=\"+mn-ea\"/>\n								<a:cs typeface=\"+mn-cs\"/>\n							</a:defRPr>\n						</a:pPr>\n						<a:r>\n							<a:rPr lang=\"en-US\" sz=\"1000\"/>\n							<a:t>" + title1 + "</a:t>\n						</a:r>\n						<a:r>\n							<a:rPr lang=\"en-US\" sz=\"1000\" baseline=\"0\"/>\n							<a:t>" + title2 + "</a:t>\n						</a:r>\n						<a:endParaRPr lang=\"en-US\" sz=\"1000\"/>\n					</a:p>\n				</c:rich>\n			</c:tx>\n			<c:overlay val=\"0\"/>\n			<c:spPr>\n				<a:noFill/>\n				<a:ln>\n					<a:noFill/>\n				</a:ln>\n				<a:effectLst/>\n			</c:spPr>\n			<c:txPr>\n				<a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n				<a:lstStyle/>\n				<a:p>\n					<a:pPr>\n						<a:defRPr sz=\"1000\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" spc=\"0\" baseline=\"0\">\n							<a:solidFill>\n								<a:schemeClr val=\"tx1\">\n									<a:lumMod val=\"65000\"/>\n									<a:lumOff val=\"35000\"/>\n								</a:schemeClr>\n							</a:solidFill>\n							<a:latin typeface=\"+mn-lt\"/>\n							<a:ea typeface=\"+mn-ea\"/>\n							<a:cs typeface=\"+mn-cs\"/>\n						</a:defRPr>\n					</a:pPr>\n					<a:endParaRPr lang=\"en-US\"/>\n				</a:p>\n			</c:txPr>\n		</c:title>\n		<c:autoTitleDeleted val=\"0\"/>\n		<c:plotArea>\n			<c:layout/>\n			<c:barChart>\n				<c:barDir val=\"col\"/>\n				<c:grouping val=\"clustered\"/>\n				<c:varyColors val=\"0\"/>";
      case 'line':
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<c:chartSpace\n	xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"\n	xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"\n	xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"\n	xmlns:c16r2=\"http://schemas.microsoft.com/office/drawing/2015/06/chart\">\n	<c:date1904 val=\"0\"/>\n	<c:lang val=\"en-US\"/>\n	<c:roundedCorners val=\"0\"/>\n	<mc:AlternateContent\n		xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">\n		<mc:Choice Requires=\"c14\"\n			xmlns:c14=\"http://schemas.microsoft.com/office/drawing/2007/8/2/chart\">\n			<c14:style val=\"102\"/>\n		</mc:Choice>\n		<mc:Fallback>\n			<c:style val=\"2\"/>\n		</mc:Fallback>\n	</mc:AlternateContent>\n	<c:chart>\n		<c:title>\n			<c:tx>\n				<c:rich>\n					<a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n					<a:lstStyle/>\n					<a:p>\n						<a:pPr>\n							<a:defRPr sz=\"1000\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" spc=\"0\" baseline=\"0\">\n								<a:solidFill>\n									<a:schemeClr val=\"tx1\">\n										<a:lumMod val=\"65000\"/>\n										<a:lumOff val=\"35000\"/>\n									</a:schemeClr>\n								</a:solidFill>\n								<a:latin typeface=\"+mn-lt\"/>\n								<a:ea typeface=\"+mn-ea\"/>\n								<a:cs typeface=\"+mn-cs\"/>\n							</a:defRPr>\n						</a:pPr>\n						<a:r>\n							<a:rPr lang=\"en-US\" sz=\"1000\"/>\n							<a:t>" + title1 + "</a:t>\n						</a:r>\n						<a:r>\n							<a:rPr lang=\"en-US\" sz=\"1000\" baseline=\"0\"/>\n							<a:t>" + title2 + "</a:t>\n						</a:r>\n						<a:endParaRPr lang=\"en-US\" sz=\"1000\"/>\n					</a:p>\n				</c:rich>\n			</c:tx>\n			<c:overlay val=\"0\"/>\n			<c:spPr>\n				<a:noFill/>\n				<a:ln>\n					<a:noFill/>\n				</a:ln>\n				<a:effectLst/>\n			</c:spPr>\n			<c:txPr>\n				<a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n				<a:lstStyle/>\n				<a:p>\n					<a:pPr>\n						<a:defRPr sz=\"1000\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" spc=\"0\" baseline=\"0\">\n							<a:solidFill>\n								<a:schemeClr val=\"tx1\">\n									<a:lumMod val=\"65000\"/>\n									<a:lumOff val=\"35000\"/>\n								</a:schemeClr>\n							</a:solidFill>\n							<a:latin typeface=\"+mn-lt\"/>\n							<a:ea typeface=\"+mn-ea\"/>\n							<a:cs typeface=\"+mn-cs\"/>\n						</a:defRPr>\n					</a:pPr>\n					<a:endParaRPr lang=\"en-US\"/>\n				</a:p>\n			</c:txPr>\n		</c:title>\n		<c:autoTitleDeleted val=\"0\"/>\n		<c:plotArea>\n			<c:layout/>\n			<c:barChart>\n				<c:barDir val=\"col\"/>\n				<c:grouping val=\"clustered\"/>\n				<c:varyColors val=\"0\"/>";
    }
  };

  ChartMaker.prototype.getFormatCode = function() {
    if (this.options.axis.x.type === 'date') {
      return "<c:formatCode>m/d/yyyy</c:formatCode>";
    } else {
      return "";
    }
  };

  ChartMaker.prototype.getLineTemplate = function(chartType, line, lineCounter) {
    var elem, i, j, k, l, len, len1, len2, len3, len4, len5, m, n, o, ref, ref1, ref2, ref3, ref4, ref5, result;
    switch (chartType) {
      case 'radar':
        result = "<c:ser>\n	<c:idx val=\"" + lineCounter + "\"/>\n	<c:order val=\"" + lineCounter + "\"/>\n	<c:tx>\n		<c:v>" + line.name + "</c:v>\n	</c:tx>";
        if (lineCounter === 1) {
          result += "<c:spPr><a:ln w=\"28800\"><a:solidFill><a:srgbClr val=\"990000\"/></a:solidFill><a:prstDash val=\"dash\"/><a:round/></a:ln></c:spPr>\n";
        }
        result += "<c:marker>\n	<c:symbol val=\"none\"/>\n</c:marker>\n<c:cat>\n	<c:" + this.ref + ">\n		<c:" + this.cache + ">\n			" + (this.getFormatCode()) + "\n			<c:ptCount val=\"" + line.data.length + "\"/>\n";
        ref = line.data;
        for (i = j = 0, len = ref.length; j < len; i = ++j) {
          elem = ref[i];
          result += "<c:pt idx=\"" + i + "\">\n	<c:v>" + elem.x + "</c:v>\n</c:pt>";
        }
        result += "		</c:" + this.cache + ">\n	</c:" + this.ref + ">\n</c:cat>\n<c:val>\n	<c:numRef>\n		<c:numCache>\n			<c:formatCode>General</c:formatCode>\n			<c:ptCount val=\"" + line.data.length + "\"/>";
        ref1 = line.data;
        for (i = k = 0, len1 = ref1.length; k < len1; i = ++k) {
          elem = ref1[i];
          result += "<c:pt idx=\"" + i + "\">\n	<c:v>" + elem.y + "</c:v>\n</c:pt>";
        }
        result += "			</c:numCache>\n		</c:numRef>\n	</c:val>\n</c:ser>";
        break;
      case 'bar':
        result = "<c:ser>\n	<c:idx val=\"" + lineCounter + "\"/>\n	<c:order val=\"" + lineCounter + "\"/>\n	<c:tx>\n		<c:v>" + line.name + "</c:v>\n	</c:tx>";
        result += "<c:spPr>\n	<a:solidFill>\n		<a:schemeClr val=\"accent" + (lineCounter + 1) + "\"/>\n	</a:solidFill>\n	<a:ln>\n		<a:noFill/>\n	</a:ln>\n	<a:effectLst/>\n</c:spPr>\n<c:invertIfNegative val=\"0\"/>\n<c:dLbls>\n	<c:spPr>\n		<a:noFill/>\n		<a:ln>\n			<a:noFill/>\n		</a:ln>\n		<a:effectLst/>\n	</c:spPr>\n	<c:txPr>\n		<a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" lIns=\"38100\" tIns=\"19050\" rIns=\"38100\" bIns=\"19050\" anchor=\"ctr\" anchorCtr=\"1\">\n			<a:spAutoFit/>\n		</a:bodyPr>\n		<a:lstStyle/>\n		<a:p>\n			<a:pPr>\n				<a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n					<a:solidFill>\n						<a:schemeClr val=\"tx1\">\n							<a:lumMod val=\"75000\"/>\n							<a:lumOff val=\"25000\"/>\n						</a:schemeClr>\n					</a:solidFill>\n					<a:latin typeface=\"+mn-lt\"/>\n					<a:ea typeface=\"+mn-ea\"/>\n					<a:cs typeface=\"+mn-cs\"/>\n				</a:defRPr>\n			</a:pPr>\n			<a:endParaRPr lang=\"en-US\"/>\n		</a:p>\n	</c:txPr>\n	<c:dLblPos val=\"outEnd\"/>\n	<c:showLegendKey val=\"0\"/>\n	<c:showVal val=\"1\"/>\n	<c:showCatName val=\"0\"/>\n	<c:showSerName val=\"0\"/>\n	<c:showPercent val=\"0\"/>\n	<c:showBubbleSize val=\"0\"/>\n	<c:showLeaderLines val=\"0\"/>\n	<c:extLst>\n		<c:ext uri=\"{CE6537A1-D6FC-4f65-9D91-7224C49458BB}\"\n			xmlns:c15=\"http://schemas.microsoft.com/office/drawing/2012/chart\">\n			<c15:showLeaderLines val=\"1\"/>\n			<c15:leaderLines>\n				<c:spPr>\n					<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">\n						<a:solidFill>\n							<a:schemeClr val=\"tx1\">\n								<a:lumMod val=\"35000\"/>\n								<a:lumOff val=\"65000\"/>\n							</a:schemeClr>\n						</a:solidFill>\n						<a:round/>\n					</a:ln>\n					<a:effectLst/>\n				</c:spPr>\n			</c15:leaderLines>\n		</c:ext>\n	</c:extLst>\n</c:dLbls>\n<c:cat>\n	<c:" + this.ref + ">\n		<c:" + this.cache + ">\n			" + (this.getFormatCode()) + "\n			<c:ptCount val=\"" + line.data.length + "\"/>\n";
        ref2 = line.data;
        for (i = l = 0, len2 = ref2.length; l < len2; i = ++l) {
          elem = ref2[i];
          result += "<c:pt idx=\"" + i + "\">\n	<c:v>" + elem.x + "</c:v>\n</c:pt>";
        }
        result += "		</c:" + this.cache + ">\n	</c:" + this.ref + ">\n</c:cat>\n<c:val>\n	<c:numRef>\n		<c:numCache>\n			<c:formatCode>General</c:formatCode>\n			<c:ptCount val=\"" + line.data.length + "\"/>";
        ref3 = line.data;
        for (i = m = 0, len3 = ref3.length; m < len3; i = ++m) {
          elem = ref3[i];
          result += "<c:pt idx=\"" + i + "\">\n	<c:v>" + elem.y + "</c:v>\n</c:pt>";
        }
        result += "			</c:numCache>\n		</c:numRef>\n	</c:val>\n	<c:extLst>\n		<c:ext uri=\"{C3380CC4-5D6E-409C-BE32-E72D297353CC}\"\n			xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\">\n			<c16:uniqueId val=\"{00000000-9819-4A27-9D24-58E0D8F9109B}\"/>\n		</c:ext>\n	</c:extLst>\n</c:ser>";
        break;
      case 'line':
        result = "<c:ser>\n	<c:idx val=\"" + lineCounter + "\"/>\n	<c:order val=\"" + lineCounter + "\"/>\n	<c:tx>\n		<c:v>" + line.name + "</c:v>\n	</c:tx>";
        result += "<c:spPr>\n	<a:ln w=\"28575\" cap=\"rnd\">\n		<a:solidFill>\n			<a:schemeClr val=\"accent1\"/>\n		</a:solidFill>\n		<a:round/>\n	</a:ln>\n	<a:effectLst/>\n</c:spPr>\n<c:marker>\n	<c:symbol val=\"none\"/>\n</c:marker>\n<c:cat>\n	<c:numRef>\n		<c:numCache>\n				<c:formatCode>General</c:formatCode>\n			<c:ptCount val=\"" + line.data.length + "\"/>\n";
        ref4 = line.data;
        for (i = n = 0, len4 = ref4.length; n < len4; i = ++n) {
          elem = ref4[i];
          result += "<c:pt idx=\"" + i + "\">\n	<c:v>" + elem.x + "</c:v>\n</c:pt>";
        }
        result += "		</c:numCache>\n	</c:numRef>\n</c:cat>\n<c:val>\n	<c:numRef>\n		<c:numCache>\n			<c:formatCode>General</c:formatCode>\n			<c:ptCount val=\"" + line.data.length + "\"/>";
        ref5 = line.data;
        for (i = o = 0, len5 = ref5.length; o < len5; i = ++o) {
          elem = ref5[i];
          result += "<c:pt idx=\"" + i + "\">\n	<c:v>" + elem.y + "</c:v>\n</c:pt>";
        }
        result += "			</c:numCache>\n		</c:numRef>\n	</c:val>\n	<c:extLst>\n		<c:ext uri=\"{C3380CC4-5D6E-409C-BE32-E72D297353CC}\"\n			xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\">\n			<c16:uniqueId val=\"{00000000-9819-4A27-9D24-58E0D8F9109B}\"/>\n		</c:ext>\n	</c:extLst>\n</c:ser>";
    }
    return result;
  };

  ChartMaker.prototype.id1 = 142309248;

  ChartMaker.prototype.id2 = 142310784;

  ChartMaker.prototype.getScaling = function(opts) {
    return "<c:scaling>\n	<c:orientation val=\"" + opts.orientation + "\"/>\n	" + (opts.max !== void 0 ? "<c:max val=\"" + opts.max + "\"/>" : "") + "\n	" + (opts.min !== void 0 ? "<c:min val=\"" + opts.min + "\"/>" : "") + "\n</c:scaling>";
  };

  ChartMaker.prototype.getAxOpts = function(chartType) {
    switch (chartType) {
      case 'radar':
        return "<c:axId val=\"" + this.id1 + "\"/>\n" + (this.getScaling(this.options.axis.x)) + "\n<c:axPos val=\"b\"/>\n\n<c:delete val=\"0\"/>\n<c:majorGridlines/>\n						<c:majorTickMark val=\"cross\"/>\n						<c:minorTickMark val=\"cross\"/>\n<c:tickLblPos val=\"nextTo\"/>\n<c:txPr>\n	<a:bodyPr/>\n	<a:lstStyle/>\n	<a:p>\n		<a:pPr>\n			<a:defRPr sz=\"800\"/>\n		</a:pPr>\n		<a:endParaRPr lang=\"ru-RU\"/>\n	</a:p>\n</c:txPr>\n<c:crossAx val=\"" + this.id2 + "\"/>\n<c:crosses val=\"autoZero\"/>\n<c:auto val=\"1\"/>\n<c:lblOffset val=\"100\"/>\n<c:noMultiLvlLbl val=\"1\"/>";
      case 'bar':
        return "<c:axId val=\"" + this.id1 + "\"/>\n" + (this.getScaling(this.options.axis.x)) + "\n<c:axPos val=\"b\"/>\n\n<c:delete val=\"0\"/>\n<c:majorTickMark val=\"none\"/>\n<c:minorTickMark val=\"none\"/>\n<c:tickLblPos val=\"none\"/>\n<c:spPr>\n	<a:noFill/>\n	<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">\n		<a:solidFill>\n			<a:schemeClr val=\"tx1\">\n				<a:lumMod val=\"15000\"/>\n				<a:lumOff val=\"85000\"/>\n			</a:schemeClr>\n		</a:solidFill>\n		<a:round/>\n	</a:ln>\n	<a:effectLst/>\n</c:spPr>\n<c:txPr>\n	<a:bodyPr rot=\"-60000000\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n	<a:lstStyle/>\n	<a:p>\n		<a:pPr>\n			<a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n				<a:solidFill>\n					<a:schemeClr val=\"tx1\">\n						<a:lumMod val=\"65000\"/>\n						<a:lumOff val=\"35000\"/>\n					</a:schemeClr>\n				</a:solidFill>\n				<a:latin typeface=\"+mn-lt\"/>\n				<a:ea typeface=\"+mn-ea\"/>\n				<a:cs typeface=\"+mn-cs\"/>\n			</a:defRPr>\n		</a:pPr>\n		<a:endParaRPr lang=\"en-US\"/>\n	</a:p>\n</c:txPr>\n<c:crossAx val=\"" + this.id2 + "\"/>\n<c:crosses val=\"autoZero\"/>\n<c:auto val=\"1\"/>\n<c:lblOffset val=\"100\"/>\n<c:noMultiLvlLbl val=\"1\"/>";
      case 'line':
        return "	<c:axId val=\"" + this.id1 + "\"/>\n	" + (this.getScaling(this.options.axis.x)) + "\n	<c:axPos val=\"b\"/>\n\n<c:delete val=\"0\"/>\n				<c:axPos val=\"b\"/>\n				<c:majorTickMark val=\"out\"/>\n				<c:minorTickMark val=\"none\"/>\n				<c:tickLblPos val=\"nextTo\"/>\n				<c:spPr>\n					<a:noFill/>\n					<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">\n						<a:solidFill>\n							<a:schemeClr val=\"tx1\">\n								<a:lumMod val=\"15000\"/>\n								<a:lumOff val=\"85000\"/>\n							</a:schemeClr>\n						</a:solidFill>\n						<a:round/>\n					</a:ln>\n					<a:effectLst/>\n				</c:spPr>\n				<c:txPr>\n					<a:bodyPr rot=\"-60000000\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n					<a:lstStyle/>\n					<a:p>\n						<a:pPr>\n							<a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n								<a:solidFill>\n									<a:schemeClr val=\"tx1\"/>\n								</a:solidFill>\n								<a:latin typeface=\"+mn-lt\"/>\n								<a:ea typeface=\"+mn-ea\"/>\n								<a:cs typeface=\"+mn-cs\"/>\n							</a:defRPr>\n						</a:pPr>\n						<a:endParaRPr lang=\"en-US\"/>\n					</a:p>\n				</c:txPr>\n	<c:crossAx val=\"" + this.id2 + "\"/>\n	<c:crosses val=\"autoZero\"/>\n	<c:auto val=\"1\"/>\n	<c:lblOffset val=\"100\"/>\n	<c:noMultiLvlLbl val=\"0\"/>";
    }
  };

  ChartMaker.prototype.getCatAx = function(chartType) {
    switch (chartType) {
      case 'radar':
        return "<c:catAx>\n	" + (this.getAxOpts(chartType)) + "\n	<c:lblAlgn val=\"ctr\"/>\n</c:catAx>";
      case 'bar':
        return "<c:catAx>\n	" + (this.getAxOpts(chartType)) + "\n	<c:lblAlgn val=\"ctr\"/>\n</c:catAx>";
      case 'line':
        return "<c:catAx>\n	" + (this.getAxOpts(chartType)) + "\n	<c:lblAlgn val=\"ctr\"/>\n</c:catAx>";
    }
  };

  ChartMaker.prototype.getDateAx = function(chartType) {
    return "<c:dateAx>\n	" + (this.getAxOpts(chartType)) + "\n	<c:delete val=\"0\"/>\n	<c:numFmt formatCode=\"" + this.options.axis.x.date.code + "\" sourceLinked=\"0\"/>\n	<c:majorTickMark val=\"out\"/>\n	<c:minorTickMark val=\"none\"/>\n	<c:baseTimeUnit val=\"days\"/>\n	<c:majorUnit val=\"" + this.options.axis.x.date.step + "\"/>\n	<c:majorTimeUnit val=\"" + this.options.axis.x.date.unit + "\"/>\n</c:dateAx>";
  };

  ChartMaker.prototype.getBorder = function() {
    if (!this.options.border) {
      return "<c:spPr>\n	<a:ln>\n		<a:noFill/>\n	</a:ln>\n</c:spPr>";
    } else {
      return '';
    }
  };

  ChartMaker.prototype.getTemplateBottom = function(chartType) {
    var result;
    switch (chartType) {
      case 'radar':
        result = "	<c:marker val=\"1\"/>\n		<c:dLbls>\n				<c:showLegendKey val=\"0\"/>\n				<c:showVal val=\"0\"/>\n				<c:showCatName val=\"0\"/>\n				<c:showSerName val=\"0\"/>\n				<c:showPercent val=\"0\"/>\n				<c:showBubbleSize val=\"0\"/>\n		</c:dLbls>\n	<c:axId val=\"" + this.id1 + "\"/>\n	<c:axId val=\"" + this.id2 + "\"/>\n</c:radarChart>";
        switch (this.options.axis.x.type) {
          case 'date':
            result += this.getDateAx(chartType);
            break;
          default:
            result += this.getCatAx(chartType);
        }
        result += "			<c:valAx>\n				<c:axId val=\"" + this.id2 + "\"/>\n				" + (this.getScaling(this.options.axis.y)) + "\n				<c:axPos val=\"l\"/>\n				<c:delete val=\"0\"/>\n				" + (this.options.grid ? "<c:majorGridlines> <c:spPr> <a:ln> <a:solidFill> <a:schemeClr val=\"bg1\"> <a:lumMod val=\"85000\"/> </a:schemeClr> \</a:solidFill> \</a:ln> \</c:spPr> \</c:majorGridlines>" : "") + "\n<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>                <c:majorTickMark val=\"none\"/>\n<c:minorTickMark val=\"none\"/>\n<c:tickLblPos val=\"none\"/>\n<c:spPr>\n<a:ln>\n<a:solidFill>\n<a:schemeClr val=\"bg1\">\n<a:lumMod val=\"75000\"/>\n</a:schemeClr>\n</a:solidFill>\n</a:ln>\n</c:spPr>\n				<c:txPr>\n					<a:bodyPr/>\n					<a:lstStyle/>\n					<a:p>\n						<a:pPr>\n							<a:defRPr sz=\"600\"/>\n						</a:pPr>\n						<a:endParaRPr lang=\"ru-RU\"/>\n					</a:p>\n				</c:txPr>\n				<c:crossAx val=\"" + this.id1 + "\"/>\n				<c:crosses val=\"autoZero\"/>\n				<c:crossBetween val=\"between\"/>                <c:majorUnit val=\"1\"/>\n			</c:valAx>\n		</c:plotArea>\n		<c:legend>\n			<c:legendPos val=\"" + this.options.legend.position + "\"/>\n			<c:layout/>\n		</c:legend>\n		<c:plotVisOnly val=\"1\"/>\n	</c:chart>\n	" + (this.getBorder()) + "\n</c:chartSpace>";
        break;
      case 'bar':
        result = "	<c:dLbls>\n		<c:dLblPos val=\"outEnd\"/>\n		<c:showLegendKey val=\"0\"/>\n		<c:showVal val=\"1\"/>\n		<c:showCatName val=\"0\"/>\n		<c:showSerName val=\"0\"/>\n		<c:showPercent val=\"0\"/>\n		<c:showBubbleSize val=\"0\"/>\n	</c:dLbls>\n	<c:gapWidth val=\"219\"/>\n	<c:overlap val=\"-27\"/>\n	<c:axId val=\"" + this.id1 + "\"/>\n	<c:axId val=\"" + this.id2 + "\"/>\n</c:barChart>";
        switch (this.options.axis.x.type) {
          case 'date':
            result += this.getDateAx(chartType);
            break;
          default:
            result += this.getCatAx(chartType);
        }
        result += "				<c:valAx>\n					<c:axId val=\"" + this.id2 + "\"/>\n					" + (this.getScaling(this.options.axis.y)) + "\n					<c:axPos val=\"l\"/>\n					<c:delete val=\"0\"/>\n					<c:majorGridlines>\n						<c:spPr>\n							<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">\n								<a:noFill/>\n								<a:round/>\n							</a:ln>\n							<a:effectLst/>\n						</c:spPr>\n					</c:majorGridlines>\n					<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>\n					<c:majorTickMark val=\"none\"/>\n					<c:minorTickMark val=\"none\"/>\n					<c:tickLblPos val=\"nextTo\"/>\n					<c:spPr>\n						<a:noFill/>\n						<a:ln>\n							<a:noFill/>\n						</a:ln>\n						<a:effectLst/>\n					</c:spPr>\n					<c:txPr>\n						<a:bodyPr rot=\"-60000000\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n						<a:lstStyle/>\n						<a:p>\n							<a:pPr>\n								<a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n									<a:solidFill>\n										<a:schemeClr val=\"tx1\">\n											<a:lumMod val=\"65000\"/>\n											<a:lumOff val=\"35000\"/>\n										</a:schemeClr>\n									</a:solidFill>\n									<a:latin typeface=\"+mn-lt\"/>\n									<a:ea typeface=\"+mn-ea\"/>\n									<a:cs typeface=\"+mn-cs\"/>\n								</a:defRPr>\n							</a:pPr>\n							<a:endParaRPr lang=\"en-US\"/>\n						</a:p>\n					</c:txPr>\n					<c:crossAx val=\"" + this.id1 + "\"/>\n					<c:crosses val=\"autoZero\"/>\n					<c:crossBetween val=\"between\"/>\n				</c:valAx>\n				<c:spPr>\n					<a:noFill/>\n					<a:ln>\n						<a:noFill/>\n					</a:ln>\n					<a:effectLst/>\n				</c:spPr>\n			</c:plotArea>\n			<c:legend>\n				<c:overlay val=\"0\"/>\n				<c:spPr>\n					<a:noFill/>\n					<a:ln>\n						<a:noFill/>\n					</a:ln>\n					<a:effectLst/>\n				</c:spPr>\n				<c:txPr>\n					<a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n					<a:lstStyle/>\n					<a:p>\n						<a:pPr>\n							<a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n								<a:solidFill>\n									<a:schemeClr val=\"tx1\">\n										<a:lumMod val=\"65000\"/>\n										<a:lumOff val=\"35000\"/>\n									</a:schemeClr>\n								</a:solidFill>\n								<a:latin typeface=\"+mn-lt\"/>\n								<a:ea typeface=\"+mn-ea\"/>\n								<a:cs typeface=\"+mn-cs\"/>\n							</a:defRPr>\n						</a:pPr>\n						<a:endParaRPr lang=\"en-US\"/>\n					</a:p>\n				</c:txPr>\n			<c:legendPos val=\"" + this.options.legend.position + "\"/>\n		</c:legend>\n		<c:plotVisOnly val=\"1\"/>\n		<c:dispBlanksAs val=\"gap\"/>\n		<c:extLst>\n			<c:ext uri=\"{56B9EC1D-385E-4148-901F-78D8002777C0}\"\n				xmlns:c16r3=\"http://schemas.microsoft.com/office/drawing/2017/03/chart\">\n				<c16r3:dataDisplayOptions16>\n					<c16r3:dispNaAsBlank val=\"1\"/>\n				</c16r3:dataDisplayOptions16>\n			</c:ext>\n		</c:extLst>\n		<c:showDLblsOverMax val=\"0\"/>\n	</c:chart>\n	<c:spPr>\n		<a:solidFill>\n			<a:schemeClr val=\"bg1\"/>\n		</a:solidFill>\n		<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">\n			<a:noFill/>\n			<a:round/>\n		</a:ln>\n		<a:effectLst/>\n	</c:spPr>\n	<c:txPr>\n		<a:bodyPr/>\n		<a:lstStyle/>\n		<a:p>\n			<a:pPr>\n				<a:defRPr/>\n			</a:pPr>\n			<a:endParaRPr lang=\"en-US\"/>\n		</a:p>\n	</c:txPr>\n</c:chartSpace>";
        break;
      case 'line':
        result = "	<c:dLbls>\n		<c:showLegendKey val=\"0\"/>\n		<c:showVal val=\"0\"/>\n		<c:showCatName val=\"0\"/>\n		<c:showSerName val=\"0\"/>\n		<c:showPercent val=\"0\"/>\n		<c:showBubbleSize val=\"0\"/>\n	</c:dLbls>\n	<c:smooth val=\"0\"/>\n	<c:axId val=\"" + this.id1 + "\"/>\n	<c:axId val=\"" + this.id2 + "\"/>\n</c:barChart>";
        switch (this.options.axis.x.type) {
          case 'date':
            result += this.getDateAx(chartType);
            break;
          default:
            result += this.getCatAx(chartType);
        }
        result += "				<c:valAx>\n					<c:axId val=\"" + this.id2 + "\"/>\n					" + (this.getScaling(this.options.axis.y)) + "\n					<c:axPos val=\"l\"/>\n					<c:delete val=\"0\"/>\n					<c:majorGridlines>\n						<c:spPr>\n							<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">\n								<a:noFill/>\n								<a:round/>\n							</a:ln>\n							<a:effectLst/>\n						</c:spPr>\n					</c:majorGridlines>\n					<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>\n					<c:majorTickMark val=\"none\"/>\n					<c:minorTickMark val=\"none\"/>\n					<c:tickLblPos val=\"nextTo\"/>\n					<c:spPr>\n						<a:noFill/>\n						<a:ln>\n							<a:noFill/>\n						</a:ln>\n						<a:effectLst/>\n					</c:spPr>\n					<c:txPr>\n						<a:bodyPr rot=\"-60000000\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n						<a:lstStyle/>\n						<a:p>\n							<a:pPr>\n								<a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n									<a:solidFill>\n										<a:schemeClr val=\"tx1\">\n											<a:lumMod val=\"65000\"/>\n											<a:lumOff val=\"35000\"/>\n										</a:schemeClr>\n									</a:solidFill>\n									<a:latin typeface=\"+mn-lt\"/>\n									<a:ea typeface=\"+mn-ea\"/>\n									<a:cs typeface=\"+mn-cs\"/>\n								</a:defRPr>\n							</a:pPr>\n							<a:endParaRPr lang=\"en-US\"/>\n						</a:p>\n					</c:txPr>\n					<c:crossAx val=\"" + this.id1 + "\"/>\n					<c:crosses val=\"autoZero\"/>\n					<c:crossBetween val=\"between\"/>\n				</c:valAx>\n				<c:spPr>\n					<a:noFill/>\n					<a:ln>\n						<a:noFill/>\n					</a:ln>\n					<a:effectLst/>\n				</c:spPr>\n			</c:plotArea>\n			<c:legend>\n				<c:overlay val=\"0\"/>\n				<c:spPr>\n					<a:noFill/>\n					<a:ln>\n						<a:noFill/>\n					</a:ln>\n					<a:effectLst/>\n				</c:spPr>\n				<c:txPr>\n					<a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n					<a:lstStyle/>\n					<a:p>\n						<a:pPr>\n							<a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n								<a:solidFill>\n									<a:schemeClr val=\"tx1\">\n										<a:lumMod val=\"65000\"/>\n										<a:lumOff val=\"35000\"/>\n									</a:schemeClr>\n								</a:solidFill>\n								<a:latin typeface=\"+mn-lt\"/>\n								<a:ea typeface=\"+mn-ea\"/>\n								<a:cs typeface=\"+mn-cs\"/>\n							</a:defRPr>\n						</a:pPr>\n						<a:endParaRPr lang=\"en-US\"/>\n					</a:p>\n				</c:txPr>\n			<c:legendPos val=\"" + this.options.legend.position + "\"/>\n		</c:legend>\n		<c:plotVisOnly val=\"1\"/>\n		<c:dispBlanksAs val=\"gap\"/>\n		<c:extLst>\n			<c:ext uri=\"{56B9EC1D-385E-4148-901F-78D8002777C0}\"\n				xmlns:c16r3=\"http://schemas.microsoft.com/office/drawing/2017/03/chart\">\n				<c16r3:dataDisplayOptions16>\n					<c16r3:dispNaAsBlank val=\"1\"/>\n				</c16r3:dataDisplayOptions16>\n			</c:ext>\n		</c:extLst>\n		<c:showDLblsOverMax val=\"0\"/>\n	</c:chart>\n	<c:spPr>\n		<a:solidFill>\n			<a:schemeClr val=\"bg1\"/>\n		</a:solidFill>\n		<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">\n			<a:noFill/>\n			<a:round/>\n		</a:ln>\n		<a:effectLst/>\n	</c:spPr>\n	<c:txPr>\n		<a:bodyPr/>\n		<a:lstStyle/>\n		<a:p>\n			<a:pPr>\n				<a:defRPr/>\n			</a:pPr>\n			<a:endParaRPr lang=\"en-US\"/>\n		</a:p>\n	</c:txPr>\n</c:chartSpace>";
    }
    return result;
  };

  function ChartMaker(zip, options) {
    this.zip = zip;
    this.options = options;
    if (this.options.axis.x.type === 'date') {
      this.ref = "numRef";
      this.cache = "numCache";
    } else {
      this.ref = "strRef";
      this.cache = "strCache";
    }
  }

  ChartMaker.prototype.makeChartFile = function(chart) {
    var i, j, len, line, ref, result;
    result = this.getTemplateTop(chart.chartType, chart.title1, chart.title2);
    ref = chart.lines;
    for (i = j = 0, len = ref.length; j < len; i = ++j) {
      line = ref[i];
      result += this.getLineTemplate(chart.chartType, line, i);
    }
    result += this.getTemplateBottom(chart.chartType);
    this.chartContent = result;
    return this.chartContent;
  };

  ChartMaker.prototype.writeFile = function(path) {
    this.zip.file("word/charts/" + path + ".xml", this.chartContent, {});
  };

  return ChartMaker;

})();
