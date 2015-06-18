var fs = require('fs');
var cheerio = require('cheerio');

var filePath = "resources/powerpoint/unzipped/ppt/slides";
var files = fs.readdirSync(filePath);

files.forEach(function(filename){

	if (filename.indexOf(".xml") == -1){
		return;
	}

	console.log(filename);

	var powerpointFile = fs.readFileSync(filePath +"/" + filename);

	// parseString(powerpointFile, function (err, result) {
	//     console.log(JSON.stringify(result, null, "  "));
	// });

	var $ = cheerio.load(powerpointFile);

	$('p\\:pic p\\:cnvpr').each(function(){

		console.log($(this).attr('descr'));

	});

});


