var fs = require('fs');
var path = require('path');

var unzip = require('unzip');
var cheerio = require('cheerio');
var minimist = require('minimist');
var jade = require('jade');

var aws = require('aws-sdk');

var AWS_ACCESS_KEY = process.env.AWS_ACCESS_KEY;
var AWS_SECRET_KEY = process.env.AWS_SECRET_KEY;
var S3_BUCKET = "joelanman-powerpoint";

aws.config.update({accessKeyId: AWS_ACCESS_KEY, secretAccessKey: AWS_SECRET_KEY});

var s3Stream = require('s3-upload-stream')(new aws.S3());
var s3 = new aws.S3();

var argv = minimist(process.argv.slice(2));

var presentationName = argv._[0];

if (!presentationName){
    console.error("Error - No presentation selected.");
    process.exit(1);
}

console.log('unzipping "resources/powerpoint/' + presentationName + '.pptx"');

var uploads = 0;

var slideRegex = /^ppt\/slides\/[^\/]*\.xml$/;
var relsRegex  = /^ppt\/slides\/_rels\/[^\/]*\.rels$/;
var mediaRegex = /^ppt\/media\/[^\/]*$/;

var slideNames = [];

// read pptx and unzip to s3

fs.createReadStream(path.join("input", presentationName + ".pptx"))
.pipe(unzip.Parse())
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
            "Key": file.path,
            "ContentType": "text/plain; charset=UTF-8"
        }).on('uploaded', function (details) {

            uploads--;
        
            if (uploads == 0){
                console.log("all uploaded");
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

    // html

    var options = {
        pretty: true
    }

    var locals = {
        slides: processedSlides,
        presentationName: presentationName
    }

    var template = fs.readFileSync(path.join("lib", "slides.jade"));

    var fn = jade.compile(template, options);
    var html = fn(locals);

    fs.writeFileSync(path.join("output", presentationName + ".html"), html);

    console.log("all done");

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

    var params = {Bucket: S3_BUCKET, Key: "ppt/slides/" + slideName};

    s3.getObject(params)
        .on('success', function(response) {
            processSlide(slideName, rels, response.data.Body);
        })
        .send();

}

function processUnzipped(){

    var slides = [];

    slideNames.forEach(function(slideName){

        var params = {Bucket: S3_BUCKET, Key: "ppt/slides/_rels/" + slideName + ".rels"};

        s3.getObject(params)
            .on('success', function(response) {
                processRels(slideName, response.data.Body);
            })
            .send();

    });

}

    // json

    // console.log(JSON.stringify(slides, null, "  "));

    // console.log();

    // csv

    // console.log("slide,id,name,description");

    // slides.forEach(function(slide){

    //  slide.pics.forEach(function(pic){

    //      console.log(slide.slide + ',' + pic.id + ',"' + pic.name.replace(/"/g,'""') + '","' + pic.description.replace(/"/g,'""') + '"');

    //  });

    //  slide.groups.forEach(function(group){

    //      console.log(slide.slide + ',' + group.id + ',"' + group.name.replace(/"/g,'""') + '","' +group.description.replace(/"/g,'""') + '"');

    //  });

    // });

    // //yaml

    // console.log(yaml.stringify(slides, 4));
