var jszip = new require("node-zip");
var parser = require("xml2json");
var fs = require("fs");

var docx = module.exports = function() {};
var seperateRules = {
    // body : {
    //     open : "<w:body>",
    //     close : "</w:body>",
    //     counter : 0
    // },
    paragraph : {
        open : "<w:p [^>]+>",
        close : "</w:p>",
        counter : 0,
        type : "paragraph"
    }
};

var removeRules = {
    wlang : {
        open : "<w:lang [^/]+/>"
    },
    pStyle : {
        open : "<w:pStyle [^/]+/>"
    },
    rFonts : {
        open : "<w:rFonts [^/]+/>"
    },
    t : {
        open : "(<w:t>|<w:t [^>]+>)",
        close : "</w:t>"
    },
    rPr : {
        open : "(<w:rPr>|<w:rPr [^>]+>)",
        close : "</w:rPr>"
    },
    pPr : {
        open : "(<w:pPr>|<w:pPr [^>]+>)",
        close : "</w:pPr>"
    },
    r : {
        open : "<w:r>",
        close : "</w:r>"
    }
};

docx.prototype.getContent = function(fn) {
    fs.readFile('./.tmp/sample-docx/word/document.xml', 'utf8', function(err, data){
        if(err) throw err;
        fn(data);
    });
};

docx.prototype.parseXML = function(data) {
    var json = {};
    for(var i in seperateRules) {
        var patt = new RegExp("(?:"+seperateRules[i].open+")(.*?)(?="+seperateRules[i].close+")","g");
        json[seperateRules[i].type] = [];
        while((match = patt.exec(data)) !== null) {
            if(seperateRules[i].type == "paragraph") {
                json[seperateRules[i].type].push(match[0].replace(new RegExp(seperateRules[i].open,""),""));
            }
        }
    }
    for(var j in json.paragraph) {
        json.paragraph[j] = this.cleanParse(json.paragraph[j]);
    }
    console.log(json.paragraph);
    return json.paragraph.join("\n");
};

docx.prototype.cleanParse = function(data) {
    var self = this;
    for(var i in removeRules) {
        if(removeRules[i].close === undefined) {
            var singlePatt = new RegExp(removeRules[i].open,"g");
            data = data.replace(singlePatt,"");
        } else {
            var patt = new RegExp("(?="+removeRules[i].open+")(.*?)(?="+removeRules[i].close+")","g");
            while((match = patt.exec(data)) !== null) {
                var temp = match[0].replace(new RegExp(removeRules[i].open,""),"");
                data = data.replace(new RegExp("(?="+removeRules[i].open+").*?"+removeRules[i].close,""),temp);
                //console.log(temp);
            }
        }
    }
    return data;
};

docx.prototype.init = function(){
    var self = this;
    this.getContent(function(data){
        //console.log(parser.toJson(data));
        var result = self.parseXML(data);
        fs.writeFile('./.tmp/result.txt',result, function(err){
            if(err) throw err;
            console.log("saved");
        });
    });
};