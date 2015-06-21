var express = require('express');
var router = express.Router();

var unzip = require('unzip');
var cheerio = require('cheerio');

var aws = require('aws-sdk');

var AWS_ACCESS_KEY = process.env.AWS_ACCESS_KEY;
var AWS_SECRET_KEY = process.env.AWS_SECRET_KEY;
var S3_BUCKET = "joelanman-powerpoint";

aws.config.update({accessKeyId: AWS_ACCESS_KEY, secretAccessKey: AWS_SECRET_KEY});

var s3Stream = require('s3-upload-stream')(new aws.S3());
var s3 = new aws.S3();

var now = require("performance-now");

router.get('/', function(req, res, next) {
  res.render('index', { title: 'Express' });
});

router.get('/sign_s3', function(req, res){

    aws.config.update({accessKeyId: AWS_ACCESS_KEY, secretAccessKey: AWS_SECRET_KEY});
    var s3 = new aws.S3();
    var s3_params = {
        Bucket: S3_BUCKET,
        Key: req.query.file_name.toLowerCase(),
        Expires: 60,
        ContentType: req.query.file_type,
        ACL: 'public-read'
    };

    s3.getSignedUrl('putObject', s3_params, function(err, data){
        if (err){
            console.log(err);
        }
        else{
            var return_data = {
                signed_request: data,
                url: 'https://'+S3_BUCKET+'.s3.amazonaws.com/'+req.query.file_name
            };
            res.write(JSON.stringify(return_data));
            res.end();
        }
    });

});

router.get('/report/:presentationName', function (req, res){

	var timings = [];

	timings.push({"init": now()});

	var presentationName = req.params.presentationName;

	if (!presentationName){
	    res.status(404).send("no presentation selected");
	}

	var uploads = 0;

	var slideRegex = /^ppt\/slides\/[^\/]*\.xml$/;
	var relsRegex  = /^ppt\/slides\/_rels\/[^\/]*\.rels$/;
	var mediaRegex = /^ppt\/media\/[^\/]*$/;

	var slideNames = [];

	// read pptx and unzip to s3

	var params = {Bucket: S3_BUCKET, Key: presentationName + ".pptx"};

	var readStream = s3.getObject(params).createReadStream().on("error", function(error){

		res.status(error.statusCode).send(error);

	});

	readStream.pipe(unzip.Parse())
		.on("error", function(error){

			res.status(500).send(error);

		})
	    .on('entry', function (file) {

	        var keep = false;

	        if (slideRegex.test(file.path)){

	            var slideName = file.path.replace("ppt/slides/","");

	            slideNames.push(slideName);
	            keep = true;

	        } else if (relsRegex.test(file.path)){

	            keep = true;

	        } else if (mediaRegex.test(file.path)){

	            keep = true;

	        }

	        if (keep){

	            uploads++;

	            file.pipe(s3Stream.upload({
	                "Bucket": S3_BUCKET,
	                "Key": presentationName.replace(".pptx","") + '/' + file.path,
	                "ContentType": "text/plain; charset=UTF-8"
	            }).on('uploaded', function (details) {

	                uploads--;
	            
	                if (uploads == 0){
	                    console.log("all uploaded");
						timings.push({"uploaded": now()});
	                    processUnzipped();
	                }

	            }));

	        } else {

	            file.autodrain();

	        }
	    })
	    .on('close', function(){

	        console.log("done unzipping");

	    });

	processedSlides = [];

	function queueSlideData(slide){

	    processedSlides.push(slide);

	    if (processedSlides.length != slideNames.length){
	        return;
	    }

	    processedSlides.sort(function(a,b){

	        if (a.slide < b.slide) return -1;
	        if (a.slide > b.slide) return 1;
	        if (a.slide == b.slide) return 0;

	    });

		timings.push({"processed": now()});

	    var locals = {
	        slides: processedSlides,
	        presentationName: presentationName
	    }

	    console.log(JSON.stringify(timings, null, '  '));

	    res.render("report", locals)

	}

	function processSlide(slideName, rels, slideData){

	    console.log("processSlide: " + slideName);

	    var $ = cheerio.load(slideData);

	    var text = "";

	    $("p\\:txBody a\\:p a\\:r a\\:t").each(function(){

	        text += $(this).text() + " ";

	    });

	    var slide = {
	        "slide": Number(slideName.replace("slide", "").replace(".xml", "")),
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

	    $('p\\:grpSp p\\:nvGrpSpPr p\\:cNvPr').each(function(){

	        var $this = $(this);

	        slide.groups.push({
	            "id": $this.attr('id'),
	            "name": $this.attr('name'),
	            "description": $this.attr('descr') || ""
	        });

	    });

	    queueSlideData(slide);

	}

	function processRels(slideName, relsData){

	    var $ = cheerio.load(relsData);

	    var rels = {};

	    $("Relationship").each(function(){

	        var $this = $(this);

	        rels[$this.attr("id")] = $this.attr("target").replace("../media/","");

	    });

	    var params = {Bucket: S3_BUCKET, Key: presentationName + "/ppt/slides/" + slideName};

	    s3.getObject(params)
	        .on('success', function(response) {
	            processSlide(slideName, rels, response.data.Body);
	        })
	        .send();

	}

	function processUnzipped(){

	    var slides = [];

	    slideNames.forEach(function(slideName){

	        var params = {Bucket: S3_BUCKET, Key: presentationName + "/ppt/slides/_rels/" + slideName + ".rels"};

	        s3.getObject(params)
	            .on('success', function(response) {
	                processRels(slideName, response.data.Body);
	            })
	            .send();

	    });
	}
});

module.exports = router;
