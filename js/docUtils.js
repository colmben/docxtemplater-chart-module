var DOMParser, DocUtils, XMLSerializer;

DOMParser = require('xmldom').DOMParser;

XMLSerializer = require('xmldom').XMLSerializer;

DocUtils = require('docxtemplater').DocUtils;

DocUtils.xml2Str = function(xmlNode) {
  var a;
  a = new XMLSerializer();
  return a.serializeToString(xmlNode);
};

DocUtils.Str2xml = function(str, errorHandler) {
  var parser, xmlDoc;
  parser = new DOMParser({
    errorHandler: errorHandler
  });
  return xmlDoc = parser.parseFromString(str, "text/xml");
};

DocUtils.maxArray = function(a) {
  return Math.max.apply(null, a);
};

module.exports = DocUtils;
