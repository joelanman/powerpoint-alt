var fs = require('fs');
var cheerio = require('cheerio');
var jade = require('jade');


var filePath = "resources/powerpoint/unzipped/ppt/slides";
var files = fs.readdirSync(filePath);

var slides = [];

files.forEach(function(filename){

	if (filename.indexOf(".xml") == -1){
		return;
	}

	// console.log(filename);

	var relsFile = fs.readFileSync(filePath +"/_rels/" + filename + ".rels");

	var $ = cheerio.load(relsFile);

	var rels = {};

	$("Relationship").each(function(){

		var $this = $(this);

		rels[$this.attr("id")] = $this.attr("target").replace("../media/","");

	});

	var slideFile = fs.readFileSync(filePath +"/" + filename);

	var $ = cheerio.load(slideFile);

	var text = "";

	$("p\\:txBody a\\:p a\\:r a\\:t").each(function(){

		text += $(this).text() + " ";

	});

	var slide = {
		"slide": Number(filename.replace("slide", "").replace(".xml", "")),
		"text": text,
		"pics": [],
		"groups": []
	};

	// pics

	$('p\\:pic p\\:cnvpr').each(function(){

		var $this = $(this);

		var relId = $this.closest("p\\:pic").find("a\\:blip").attr("r:embed");

		slide.pics.push({
			"id": $this.attr('id'),
			"name": $this.attr('name'),
			"description": $this.attr('descr') || "",
			"file": rels[relId]
		});

	});

	// groups

	//p:grpSp/p:nvGrpSpPr/p:cNvPr

	$('p\\:grpSp p\\:nvGrpSpPr p\\:cNvPr').each(function(){

		var $this = $(this);

		slide.groups.push({
			"id": $this.attr('id'),
			"name": $this.attr('name'),
			"description": $this.attr('descr') || ""
		});

	});

	slides.push(slide);

});

slides.sort(function(a,b){

	if (a.slide < b.slide) return -1;
	if (a.slide > b.slide) return 1;
	if (a.slide == b.slide) return 0;

});

// json

// console.log(JSON.stringify(slides, null, "  "));

// console.log();

var options = {
	pretty: true
}

var locals = {
	slides: slides
}

var template = fs.readFileSync('slides.jade');

var fn = jade.compile(template, options);
var html = fn(locals);

fs.writeFileSync("slides.html",html);

// csv

// console.log("slide,id,name,description");

// slides.forEach(function(slide){

// 	slide.pics.forEach(function(pic){

// 		console.log(slide.slide + ',' + pic.id + ',"' + pic.name.replace(/"/g,'""') + '","' + pic.description.replace(/"/g,'""') + '"');

// 	});

// 	slide.groups.forEach(function(group){

// 		console.log(slide.slide + ',' + group.id + ',"' + group.name.replace(/"/g,'""') + '","' +group.description.replace(/"/g,'""') + '"');

// 	});

// });

// //yaml

// console.log(yaml.stringify(slides, 4));
