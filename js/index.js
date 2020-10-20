var ChartMaker, ChartManager, ChartModule, SubContent;

SubContent = require('docxtemplater').SubContent;

ChartManager = require('./chartManager');

ChartMaker = require('./chartMaker');

ChartModule = (function() {

  /**
  	 * self name for self-identification, variable for fast changing;
  	 * @type {String}
   */
  ChartModule.prototype.name = 'chart';


  /**
  	 * initialize options with empty object if not recived
  	 * @manager = ModuleManager instance
  	 * @param  {Object} @options params for the module
   */

  function ChartModule(options1) {
    this.options = options1 != null ? options1 : {};
  }

  ChartModule.prototype.handleEvent = function(event, eventData) {
    var gen;
    if (event === 'rendering-file') {
      this.renderingFileName = eventData;
      gen = this.manager.getInstance('gen');
      this.chartManager = new ChartManager(gen.zip, this.renderingFileName);
      return this.chartManager.loadChartRels();
    } else if (event === 'rendered') {
      return this.finished();
    }
  };

  ChartModule.prototype.get = function(data) {
    var templaterState;
    if (data === 'loopType') {
      templaterState = this.manager.getInstance('templaterState');
      if (templaterState.textInsideTag[0] === '$') {
        return this.name;
      }
    }
    return null;
  };

  ChartModule.prototype.handle = function(type, data) {
    if (type === 'replaceTag' && data === this.name) {
      this.replaceTag();
    }
    return null;
  };

  ChartModule.prototype.finished = function() {};

  ChartModule.prototype.on = function(event, data) {
    if (event === 'error') {
      throw data;
    }
  };

  ChartModule.prototype.replaceBy = function(text, outsideElement) {
    var subContent, templaterState, xmlTemplater;
    console.log('In replaceBy, text, outsideElement', text, outsideElement);
    xmlTemplater = this.manager.getInstance('xmlTemplater');
    templaterState = this.manager.getInstance('templaterState');
    console.log('xmlTemplater.content :', xmlTemplater.content);
    console.log('Sub content :', new SubContent(xmlTemplater.content));
    subContent = new SubContent(xmlTemplater.content).getInnerTag(templaterState).getOuterXml(outsideElement);
    console.log('In replaceBy, subContent', subContent);
    return xmlTemplater.replaceXml(subContent, text);
  };

  ChartModule.prototype.convertPixelsToEmus = function(pixel) {
    return Math.round(pixel * 9525);
  };

  ChartModule.prototype.extendDefaults = function(options) {
    var deepMerge, defaultOptions, result;
    deepMerge = function(target, source) {
      var key, next, original;
      for (key in source) {
        original = target[key];
        next = source[key];
        if (original && next && typeof next === 'object') {
          deepMerge(original, next);
        } else {
          target[key] = next;
        }
      }
      return target;
    };
    defaultOptions = {
      width: 6477000 / 9525,
      height: 4152900 / 9525,
      grid: true,
      border: false,
      title: false,
      format: 'General',
      legend: {
        position: 'r'
      },
      axis: {
        x: {
          orientation: 'minMax',
          min: void 0,
          max: void 0,
          type: void 0,
          date: {
            format: 'unix',
            code: 'mm/yy',
            unit: 'months',
            step: '1'
          }
        },
        y: {
          orientation: 'minMax',
          min: void 0,
          max: void 0
        }
      }
    };
    result = deepMerge({}, defaultOptions);
    result = deepMerge(result, options);
    return result;
  };

  ChartModule.prototype.convertUnixTo1900 = function(chartData, axName) {
    var convertOption, data, i, j, len, len1, line, ref, ref1, unixTo1900;
    unixTo1900 = function(value) {
      return Math.round(value / 86400 + 25569);
    };
    convertOption = function(name) {
      if (chartData.options.axis[axName][name]) {
        return chartData.options.axis[axName][name] = unixTo1900(chartData.options.axis[axName][name]);
      }
    };
    convertOption('min');
    convertOption('max');
    ref = chartData.lines;
    for (i = 0, len = ref.length; i < len; i++) {
      line = ref[i];
      ref1 = line.data;
      for (j = 0, len1 = ref1.length; j < len1; j++) {
        data = ref1[j];
        data[axName] = unixTo1900(data[axName]);
      }
    }
    return chartData;
  };

  ChartModule.prototype.replaceTag = function() {
    var ax, chart, chartData, chartId, filename, gen, imageRels, name, newText, options, scopeManager, tag, tagXml, templaterState;
    scopeManager = this.manager.getInstance('scopeManager');
    templaterState = this.manager.getInstance('templaterState');
    gen = this.manager.getInstance('gen');
    tag = templaterState.textInsideTag.substr(1);
    console.log("Chart tag : ", tag);
    chartData = scopeManager.getValueFromScope(tag);
    console.log("Chart chartData : ", chartData);
    if (chartData == null) {
      return;
    }
    filename = tag + (this.chartManager.maxRid + 1);
    imageRels = this.chartManager.loadChartRels();
    if (!imageRels) {
      return;
    }
    chartId = this.chartManager.addChartRels(filename);
    options = this.extendDefaults(chartData.options);
    for (name in options.axis) {
      ax = options.axis[name];
      if (ax.type === 'date' && ax[ax.type].format === 'unix') {
        chartData = this.convertUnixTo1900(chartData, name);
      }
    }
    chart = new ChartMaker(gen.zip, options);
    chart.makeChartFile(chartData);
    chart.writeFile(filename);
    tagXml = this.manager.getInstance('xmlTemplater').fileTypeConfig.tagsXmlArray[0];
    console.log("Chart tagXmlArray : ", this.manager.getInstance('xmlTemplater').fileTypeConfig.tagsXmlArray);
    newText = this.getChartXml({
      chartID: chartId,
      width: this.convertPixelsToEmus(options.width),
      height: this.convertPixelsToEmus(options.height)
    });
    console.log('Chart newText, tagXml', newText, tagXml);
    return this.replaceBy(newText, tagXml);
  };

  ChartModule.prototype.getChartXml = function(arg) {
    var chartID, height, width;
    chartID = arg.chartID, width = arg.width, height = arg.height;
    return "<w:drawing>\n	<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">\n		<wp:extent cx=\"" + width + "\" cy=\"" + height + "\"/>\n		<wp:effectExtent b=\"0\" l=\"0\" r=\"0\" t=\"0\"/>\n		<wp:docPr id=\"1\" name=\"Диаграмма 1\"/>\n		<wp:cNvGraphicFramePr/>\n		<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n			<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">\n				<c:chart r:id=\"rId" + chartID + "\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>\n			</a:graphicData>\n		</a:graphic>\n	</wp:inline>\n</w:drawing>";
  };

  return ChartModule;

})();

module.exports = ChartModule;
