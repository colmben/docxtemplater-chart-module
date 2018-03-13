var ChartMaker, DocUtils;

DocUtils = require('./docUtils');

module.exports = ChartMaker = (function() {
  ChartMaker.prototype.getTemplateTop = function(chartType) {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<c:chartSpace xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n	<c:lang val=\"ru-RU\"/>\n	<c:chart>\n		" + (this.options.title ? "" : "<c:autoTitleDeleted val=\"1\"/>") + "\n		<c:plotArea>\n			<c:layout/>\n			<c:radarChart>\n			<c:radarStyle val=\"marker\"/>\n			<c:varyColors val=\"1\"/>";
  };

  ChartMaker.prototype.getFormatCode = function() {
    if (this.options.axis.x.type === 'date') {
      return "<c:formatCode>m/d/yyyy</c:formatCode>";
    } else {
      return "";
    }
  };

  ChartMaker.prototype.getLineTemplate = function(chartType, line, i) {
    var elem, j, k, len, len1, ref, ref1, result;
    result = "<c:ser>\n	<c:idx val=\"" + i + "\"/>\n	<c:order val=\"" + i + "\"/>\n	<c:tx>\n		<c:v>" + line.name + "</c:v>\n	</c:tx>";
    if (i === 1) {
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
    return result;
  };

  ChartMaker.prototype.id1 = 142309248;

  ChartMaker.prototype.id2 = 142310784;

  ChartMaker.prototype.getScaling = function(opts) {
    return "<c:scaling>\n	<c:orientation val=\"" + opts.orientation + "\"/>\n	" + (opts.max !== void 0 ? "<c:max val=\"" + opts.max + "\"/>" : "") + "\n	" + (opts.min !== void 0 ? "<c:min val=\"" + opts.min + "\"/>" : "") + "\n</c:scaling>";
  };

  ChartMaker.prototype.getAxOpts = function() {
    return "<c:axId val=\"" + this.id1 + "\"/>\n" + (this.getScaling(this.options.axis.x)) + "\n<c:axPos val=\"b\"/>\n\n<c:delete val=\"0\"/>\n    <c:majorGridlines/>\n                <c:majorTickMark val=\"cross\"/>\n                <c:minorTickMark val=\"cross\"/>\n<c:tickLblPos val=\"nextTo\"/>\n<c:txPr>\n	<a:bodyPr/>\n	<a:lstStyle/>\n	<a:p>\n		<a:pPr>\n			<a:defRPr sz=\"800\"/>\n		</a:pPr>\n		<a:endParaRPr lang=\"ru-RU\"/>\n	</a:p>\n</c:txPr>\n<c:crossAx val=\"" + this.id2 + "\"/>\n<c:crosses val=\"autoZero\"/>\n<c:auto val=\"1\"/>\n<c:lblOffset val=\"100\"/>\n    <c:noMultiLvlLbl val=\"1\"/>";
  };

  ChartMaker.prototype.getCatAx = function() {
    return "<c:catAx>\n	" + (this.getAxOpts()) + "\n	<c:lblAlgn val=\"ctr\"/>\n</c:catAx>";
  };

  ChartMaker.prototype.getDateAx = function() {
    return "<c:dateAx>\n	" + (this.getAxOpts()) + "\n	<c:delete val=\"0\"/>\n	<c:numFmt formatCode=\"" + this.options.axis.x.date.code + "\" sourceLinked=\"0\"/>\n	<c:majorTickMark val=\"out\"/>\n	<c:minorTickMark val=\"none\"/>\n	<c:baseTimeUnit val=\"days\"/>\n	<c:majorUnit val=\"" + this.options.axis.x.date.step + "\"/>\n	<c:majorTimeUnit val=\"" + this.options.axis.x.date.unit + "\"/>\n</c:dateAx>";
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
    result = "	<c:marker val=\"1\"/>\n                <c:dLbls>\n                    <c:showLegendKey val=\"0\"/>\n                    <c:showVal val=\"0\"/>\n                    <c:showCatName val=\"0\"/>\n                    <c:showSerName val=\"0\"/>\n                    <c:showPercent val=\"0\"/>\n                    <c:showBubbleSize val=\"0\"/>\n                </c:dLbls>\n	<c:axId val=\"" + this.id1 + "\"/>\n	<c:axId val=\"" + this.id2 + "\"/>\n</c:radarChart>";
    switch (this.options.axis.x.type) {
      case 'date':
        result += this.getDateAx();
        break;
      default:
        result += this.getCatAx();
    }
    result += "			<c:valAx>\n				<c:axId val=\"" + this.id2 + "\"/>\n				" + (this.getScaling(this.options.axis.y)) + "\n				<c:axPos val=\"l\"/>\n				<c:delete val=\"0\"/>\n				" + (this.options.grid ? "<c:majorGridlines> <c:spPr> <a:ln> <a:solidFill> <a:schemeClr val=\"bg1\"> <a:lumMod val=\"85000\"/> </a:schemeClr> \</a:solidFill> \</a:ln> \</c:spPr> \</c:majorGridlines>" : "") + "\n<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>                <c:majorTickMark val=\"none\"/>\n<c:minorTickMark val=\"none\"/>\n<c:tickLblPos val=\"none\"/>\n<c:spPr>\n<a:ln>\n<a:solidFill>\n<a:schemeClr val=\"bg1\">\n<a:lumMod val=\"75000\"/>\n</a:schemeClr>\n</a:solidFill>\n</a:ln>\n</c:spPr>\n				<c:txPr>\n					<a:bodyPr/>\n					<a:lstStyle/>\n					<a:p>\n						<a:pPr>\n							<a:defRPr sz=\"600\"/>\n						</a:pPr>\n						<a:endParaRPr lang=\"ru-RU\"/>\n					</a:p>\n				</c:txPr>\n				<c:crossAx val=\"" + this.id1 + "\"/>\n				<c:crosses val=\"autoZero\"/>\n				<c:crossBetween val=\"between\"/>                <c:majorUnit val=\"1\"/>\n			</c:valAx>\n		</c:plotArea>\n		<c:legend>\n			<c:legendPos val=\"" + this.options.legend.position + "\"/>\n			<c:layout/>\n		</c:legend>\n		<c:plotVisOnly val=\"1\"/>\n	</c:chart>\n	" + (this.getBorder()) + "\n</c:chartSpace>";
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

  ChartMaker.prototype.makeChartFile = function(chartType, lines) {
    var i, j, len, line, result;
    result = this.getTemplateTop(chartType);
    for (i = j = 0, len = lines.length; j < len; i = ++j) {
      line = lines[i];
      result += this.getLineTemplate(chartType, line, i);
    }
    result += this.getTemplateBottom(chartType);
    this.chartContent = result;
    return this.chartContent;
  };

  ChartMaker.prototype.writeFile = function(path) {
    this.zip.file("word/charts/" + path + ".xml", this.chartContent, {});
  };

  return ChartMaker;

})();
