var fs = require('fs');
var cheerio = require('cheerio');

var filePath = "resources/powerpoint/unzipped/ppt/slides";
var files = fs.readdirSync(filePath);

var slidesData = [];

files.forEach(function(filename){

	if (filename.indexOf(".xml") == -1){
		return;
	}

	// console.log(filename);

	var powerpointFile = fs.readFileSync(filePath +"/" + filename);

	var $ = cheerio.load(powerpointFile);

	var slideData = {
		"slide": filename.replace("slide","").replace(".xml",""),
		"pics": []
	};


	// pics

	$('p\\:pic p\\:cnvpr').each(function(){

		var $this = $(this);

		slideData.pics.push({
			"id": $this.attr('id'),
			"name": $this.attr('name'),
			"description": $this.attr('descr')
		});

	});

	slidesData.push(slideData);

});

// json

console.log(JSON.stringify(slidesData, null, "  "));

console.log();

// csv

console.log("slide,id,name,description");

slidesData.forEach(function(slide){

	slide.pics.forEach(function(pic){

		console.log(slide.slide + "," + pic.id + "," + pic.name + "," +pic.description);

	});

});
